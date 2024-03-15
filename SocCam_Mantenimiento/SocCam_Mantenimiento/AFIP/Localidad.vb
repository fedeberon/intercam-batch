Imports System.Windows.Forms
Imports helix

Public Class Localidad
    Public Property sqle As New SQLEngine
    Public Property Id As Integer = 0
    Public Property Nombre As String = ""

    Dim ConsoleOut As New ConsoleOut

    Public Property All As New Dictionary(Of String, Integer)
    Public Property AllReverse As New Dictionary(Of Integer, String)

    Public Function LoadMe(ByVal id As Integer) As Boolean
        Try
            With sqle.Query
                .Reset()
                .TableName = TABLA_LOCALIDADES.TABLA_NOMBRE
                .AddSelectColumn(TABLA_LOCALIDADES.ALL)
                .AddSelectColumn("Departamento.Nombre")
                .AddSelectColumn("Provincia.Nombre")
                .AddFirstJoin("Departamento", TABLA_LOCALIDADES.TABLA_NOMBRE, "Departamento.ID", TABLA_LOCALIDADES.ID_DEPARTAMENTO)
                .AddNestedJoin("Provincia", "Provincia.ID", "idProvincia")

                .SimpleSearch(TABLA_LOCALIDADES.ID, SQLEngineQuery.OperatorCriteria.Igual, id)

                If .Query() Then
                    If .RecordCount = 1 Then
                        Me.Id = id
                        Dim output As String = ""

                        If Not IsDBNull(.GetQueryData(TABLA_LOCALIDADES.NOMBRE)) Then
                            output = .GetQueryData(TABLA_LOCALIDADES.NOMBRE)
                        End If
                        If Not IsDBNull(.GetQueryData(3)) Then
                            output &= ", " & .GetQueryData(3)
                        End If
                        If Not IsDBNull(.GetQueryData(4)) Then
                            output &= ", " & .GetQueryData(4)
                        End If

                        Me.Nombre = output
                        Return True
                    End If
                End If

                Return False
            End With
        Catch ex As Exception
            Console.WriteLine("Error al ejecutar la consulta: " & ex.Message)
        End Try

    End Function

    Public Function LoadMeAfipFormat(ByVal myId As Integer) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_LOCALIDADES.TABLA_NOMBRE
            .AddSelectColumn(TABLA_LOCALIDADES.ALL)
            .AddSelectColumn("Departamento.Nombre")
            .AddSelectColumn("Provincia.Nombre")
            .AddFirstJoin("Departamento", TABLA_LOCALIDADES.TABLA_NOMBRE, "Departamento.ID", TABLA_LOCALIDADES.ID_DEPARTAMENTO)
            .AddNestedJoin("Provincia", "Provincia.ID", "idProvincia")

            .SimpleSearch(TABLA_LOCALIDADES.ID, SQLEngineQuery.OperatorCriteria.Igual, Id)

            If .Query() Then
                If .RecordCount = 1 Then
                    Me.Id = Id
                    Dim output As String = ""
                    If Not IsDBNull(.GetQueryData(TABLA_LOCALIDADES.NOMBRE)) Then
                        output = .GetQueryData(TABLA_LOCALIDADES.NOMBRE)
                    End If
                    If Not IsDBNull(.GetQueryData("Provincia.Nombre")) Then
                        output &= ", " & .GetQueryData("Provincia.Nombre")
                    End If
                    Me.Nombre = output
                    Return True
                End If
            End If

            Return False
        End With
    End Function

    Public Function LoadAll(Optional ByVal autosugestion As Boolean = False, Optional ByRef sugest As AutoCompleteStringCollection = Nothing) As Boolean

        With sqle.Query
            .Reset()
            .TableName = TABLA_LOCALIDADES.TABLA_NOMBRE
            .AddSelectColumn(TABLA_LOCALIDADES.ALL)
            .AddSelectColumn("Departamento.Nombre")
            .AddSelectColumn("Provincia.Nombre")
            .AddFirstJoin("Departamento", TABLA_LOCALIDADES.TABLA_NOMBRE, "Departamento.ID", TABLA_LOCALIDADES.ID_DEPARTAMENTO)
            .AddNestedJoin("Provincia", "Provincia.ID", "idProvincia")

            .SimpleSearch(TABLA_LOCALIDADES.ID, SQLEngineQuery.OperatorCriteria.Mayor, 0)

            All.Clear()
            AllReverse.Clear()
            Dim lst As New List(Of String)
            If .Query() Then
                While .QueryRead
                    Dim output As String = ""
                    If Not IsDBNull(.GetQueryData(2)) Then
                        output = .GetQueryData(2)
                    End If
                    If Not IsDBNull(.GetQueryData(3)) Then
                        output &= ", " & .GetQueryData(3)
                    End If
                    If Not IsDBNull(.GetQueryData(4)) Then
                        output &= ", " & .GetQueryData(4)
                    End If

                    Dim idLoc As Integer = .GetQueryData(0)

                    If Not All.ContainsKey(output) Then
                        All.Add(output, idLoc)
                        lst.Add(output)
                    End If

                    AllReverse.Add(idLoc, output)
                End While
                If autosugestion Then
                    sugest.AddRange(lst.ToArray)
                End If

                Return True
            Else
                Return False
            End If

        End With
    End Function


End Class
