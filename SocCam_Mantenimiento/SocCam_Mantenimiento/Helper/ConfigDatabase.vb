Imports helix
Public Class ConfigDatabase
    'True -> Production. False -> Local.
    Public Property Production As Boolean = True
    Public Property sqle As New SQLEngine
    Public Sub New()
        sqle.RequireCredentials = False
        sqle.dbType = SQLEngine.dataBaseType.SQL_SERVER
        sqle.Path = My.Computer.Name & "\" & "SQLEXPRESS"

        If (Production) Then
            sqle.DatabaseName = "soccam"
        Else
            sqle.DatabaseName = "soccam_test"
        End If
    End Sub
End Class
