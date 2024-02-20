Imports helix
Public Class MailTemplate
    Public Property sqle As New SQLEngine
    Public Property id As Integer = 0
    Public Property nombre As String = ""
    Public Property esHTML As Boolean = False
    Public Property body As String = ""
    Public Property css As String = ""
    Public Property adjunto As String = ""
    Public Property contexto As Byte = 0
    Public Property deleted As Boolean = False
    Public Property modificado As Date = Now


    Public Function LoadMe(ByVal myID As Integer) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_MAIL.TABLA_NOMBRE
            .AddSelectColumn(TABLA_MAIL.ALL)
            .SimpleSearch(TABLA_MAIL.ID, SQLEngineQuery.OperatorCriteria.Igual, myID)
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    id = CInt(.GetQueryData(TABLA_MAIL.ID))
                    nombre = CStr(.GetQueryData(TABLA_MAIL.NOMBRE))
                    esHTML = CBool(.GetQueryData(TABLA_MAIL.ES_HTML))
                    body = CStr(.GetQueryData(TABLA_MAIL.BODY))
                    css = CStr(.GetQueryData(TABLA_MAIL.CSS))
                    adjunto = CStr(.GetQueryData(TABLA_MAIL.ADJUNTO))
                    contexto = CByte(.GetQueryData(TABLA_MAIL.CONTEXTO))
                    deleted = CBool(.GetQueryData(TABLA_MAIL.DELETED))
                    modificado = CDate(.GetQueryData(TABLA_MAIL.MODIFICADO))
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
            .TableName = TABLA_MAIL.TABLA_NOMBRE
            .AddSelectColumn(TABLA_MAIL.ALL)
            .SimpleSearch(TABLA_MAIL.DELETED, SQLEngineQuery.OperatorCriteria.Igual, False)
            Return .Query(True, dt)
        End With
    End Function



    Public Function Save(ByVal editMode As Byte) As Boolean
        Select Case editMode
            Case 0
                With sqle.Insert
                    .Reset()
                    .TableName = TABLA_MAIL.TABLA_NOMBRE
                    .AddColumnValue(TABLA_MAIL.NOMBRE, nombre)
                    .AddColumnValue(TABLA_MAIL.ES_HTML, esHTML)
                    .AddColumnValue(TABLA_MAIL.BODY, body)
                    .AddColumnValue(TABLA_MAIL.CSS, css)
                    .AddColumnValue(TABLA_MAIL.ADJUNTO, adjunto)
                    .AddColumnValue(TABLA_MAIL.CONTEXTO, contexto)
                    .AddColumnValue(TABLA_MAIL.DELETED, deleted)
                    .AddColumnValue(TABLA_MAIL.MODIFICADO, Now)
                    Dim lastID As Integer = 0
                    If .Insert(lastID) Then
                        id = lastID
                        Return True
                    Else
                        Return False
                    End If
                End With
            Case 1
                With sqle.Update
                    .Reset()
                    .TableName = TABLA_MAIL.TABLA_NOMBRE
                    .AddColumnValue(TABLA_MAIL.NOMBRE, nombre)
                    .AddColumnValue(TABLA_MAIL.ES_HTML, esHTML)
                    .AddColumnValue(TABLA_MAIL.BODY, body)
                    .AddColumnValue(TABLA_MAIL.CSS, css)
                    .AddColumnValue(TABLA_MAIL.ADJUNTO, adjunto)
                    .AddColumnValue(TABLA_MAIL.CONTEXTO, contexto)
                    .AddColumnValue(TABLA_MAIL.DELETED, deleted)
                    .AddColumnValue(TABLA_MAIL.MODIFICADO, Now)
                    Return .Update
                End With
            Case Else
                Return False
        End Select
    End Function

End Class

