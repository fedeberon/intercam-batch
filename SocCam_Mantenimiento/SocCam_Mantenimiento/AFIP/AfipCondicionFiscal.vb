Imports System.IO
Imports helix

Public Class AfipCondicionFiscal
    Public Property sqle As New SQLEngine
    Public Property searchResult As New List(Of AfipCondicionFiscal)

    Public Property Cuit As String = 0
    Public Property Denominacion As String = ""
    Public Property Ganancias As Integer = 0
    Public Property Iva As Integer = 0
    Public Property Sociedad As Boolean = False
    Public Property Empleador As Boolean = False
    Public Property Actividad As Integer = 0
    Public Property Condicion As Integer = 0

    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum


    Public Structure TABLA
        Const TABLA_NOMBRE As String = "AfipCondicionFiscal"
        Const CUIT As String = TABLA_NOMBRE & "_cuit"
        Const DENOMINACION As String = TABLA_NOMBRE & "_denominacion"
        Const GANANCIAS As String = TABLA_NOMBRE & "_ganancias"
        Const IVA As String = TABLA_NOMBRE & "_iva"
        Const SOCIEDAD As String = TABLA_NOMBRE & "_sociedad"
        Const EMPLEADOR As String = TABLA_NOMBRE & "_empleador"
        Const ACTIVIDAD As String = TABLA_NOMBRE & "_actividad"
        Const CONDICION As String = TABLA_NOMBRE & "_condicion"
        Const ALL As String = CUIT & ", " & DENOMINACION & ", " & GANANCIAS & ", " & IVA & ", " & SOCIEDAD & ", " & EMPLEADOR & ", " & ACTIVIDAD & ", " & CONDICION
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
            .SimpleSearch(TABLA.CUIT, SQLEngineQuery.OperatorCriteria.Igual, myID)
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    Cuit = (.GetQueryData(TABLA.CUIT))
                    Denominacion = CStr(.GetQueryData(TABLA.DENOMINACION))
                    Ganancias = CInt(.GetQueryData(TABLA.GANANCIAS))
                    Iva = CInt(.GetQueryData(TABLA.IVA))
                    Sociedad = CBool(.GetQueryData(TABLA.SOCIEDAD))
                    Empleador = CBool(.GetQueryData(TABLA.EMPLEADOR))
                    Actividad = CInt(.GetQueryData(TABLA.ACTIVIDAD))
                    Condicion = CInt(.GetQueryData(TABLA.CONDICION))
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End With
    End Function

    Public Function LoadMe(ByVal myCuit As String) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(TABLA.CUIT, SQLEngineQuery.OperatorCriteria.Igual, myCuit.Trim)
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    Cuit = CStr(.GetQueryData(TABLA.CUIT))
                    Denominacion = CStr(.GetQueryData(TABLA.DENOMINACION))
                    Ganancias = CInt(.GetQueryData(TABLA.GANANCIAS))
                    Iva = CInt(.GetQueryData(TABLA.IVA))
                    Sociedad = CBool(.GetQueryData(TABLA.SOCIEDAD))
                    Empleador = CBool(.GetQueryData(TABLA.EMPLEADOR))
                    Actividad = CInt(.GetQueryData(TABLA.ACTIVIDAD))
                    Condicion = CInt(.GetQueryData(TABLA.CONDICION))
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
            .SimpleSearch(TABLA.CONDICION, SQLEngineQuery.OperatorCriteria.MayorIgual, 0)
            Return .Query(True, dt)
        End With
    End Function



    Public Function Save(ByVal editMode As Guardar) As Boolean
        Select Case editMode
            Case 0
                With sqle.Insert
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.CUIT, Cuit)
                    .AddColumnValue(TABLA.DENOMINACION, Denominacion)
                    .AddColumnValue(TABLA.GANANCIAS, Ganancias)
                    .AddColumnValue(TABLA.IVA, Iva)
                    .AddColumnValue(TABLA.SOCIEDAD, Sociedad)
                    .AddColumnValue(TABLA.EMPLEADOR, Empleador)
                    .AddColumnValue(TABLA.ACTIVIDAD, Actividad)
                    .AddColumnValue(TABLA.CONDICION, Condicion)
                    Dim lastID As Integer = 0
                    If .Insert(lastID) Then
                        Cuit = lastID
                        Return True
                    Else
                        Return False
                    End If
                End With
            Case 1
                With sqle.Update
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.CUIT, Cuit)
                    .AddColumnValue(TABLA.DENOMINACION, Denominacion)
                    .AddColumnValue(TABLA.GANANCIAS, Ganancias)
                    .AddColumnValue(TABLA.IVA, Iva)
                    .AddColumnValue(TABLA.SOCIEDAD, Sociedad)
                    .AddColumnValue(TABLA.EMPLEADOR, Empleador)
                    .AddColumnValue(TABLA.ACTIVIDAD, Actividad)
                    .AddColumnValue(TABLA.CONDICION, Condicion)
                    .SimpleSearch(TABLA.CUIT, SQLEngineQuery.OperatorCriteria.Igual, Cuit)
                    Return .Update
                End With
            Case Else
                Return False
        End Select
    End Function

    Public Function Delete(Optional ByVal hard As Boolean = False) As Boolean
        With sqle.Delete
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .SimpleSearch(TABLA.CUIT, SQLEngineUpdate.OperatorCriteria.Igual, Cuit)
            Return .Delete
        End With
    End Function
    Public Function DeleteAll() As Boolean
        With sqle.Delete
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .SimpleSearch(TABLA.CONDICION, SQLEngineDelete.OperatorCriteria.MayorIgual, 0)
            Return .Delete
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
                    Dim tmp As New AfipCondicionFiscal
                    tmp.Cuit = CStr(.GetQueryData(TABLA.CUIT))
                    tmp.Denominacion = CStr(.GetQueryData(TABLA.DENOMINACION))
                    tmp.Ganancias = CInt(.GetQueryData(TABLA.GANANCIAS))
                    tmp.Iva = CInt(.GetQueryData(TABLA.IVA))
                    tmp.Sociedad = CBool(.GetQueryData(TABLA.SOCIEDAD))
                    tmp.Empleador = CBool(.GetQueryData(TABLA.EMPLEADOR))
                    tmp.Actividad = CInt(.GetQueryData(TABLA.ACTIVIDAD))
                    tmp.Condicion = CInt(.GetQueryData(TABLA.CONDICION))
                    searchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function

    Public Function ImportarPadronDB(ByVal pathPadron As String, Optional silent As Boolean = False) As Boolean
        Dim ln As String = ""
        If My.Computer.FileSystem.FileExists(pathPadron) Then
            Dim lns() As String = File.ReadAllLines(pathPadron)
            Dim i As Integer = 0
            Dim totalRecords As Integer = lns.Length
            Dim ConsoleOut As New ConsoleOut
            If Not silent Then
                ConsoleOut.Print($"- Actualizando padron AFIP")
            End If
            Do
                If Not silent Then
                    ConsoleOut.UpdateLastLine($"{ConsoleOut.ProgressBarStep} Procesando {i + 1}/{totalRecords}")
                End If
                With sqle.Insert
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.CUIT, lns(i).Substring(0, 11))
                    .AddColumnValue(TABLA.DENOMINACION, lns(i).Substring(11, 30).Trim)
                    Dim tmpGanancias As Integer = 0
                    Select Case lns(i).Substring(41, 2).Trim
                        Case "", "NI", "N"
                            tmpGanancias = 0
                        Case "AC", "S"
                            tmpGanancias = 1
                        Case "EX"
                            tmpGanancias = 2
                        Case "NA"
                            tmpGanancias = 3
                        Case "XN"
                            tmpGanancias = 4
                        Case "AN"
                            tmpGanancias = 5
                        Case "NC"
                            tmpGanancias = 6
                    End Select
                    .AddColumnValue(TABLA.GANANCIAS, tmpGanancias)


                    Dim tmpIva As Integer = 0
                    Select Case lns(i).Substring(43, 2).Trim
                        Case "", "NI", "N"
                            tmpIva = 0
                        Case "AC", "S"
                            tmpIva = 1
                        Case "EX"
                            tmpIva = 2
                        Case "NA"
                            tmpIva = 3
                        Case "XN"
                            tmpIva = 4
                        Case "AN"
                            tmpIva = 5
                        Case "NC"
                            tmpIva = 6
                    End Select
                    .AddColumnValue(TABLA.IVA, tmpIva)

                    Dim tmpMonotributo As Boolean = False
                    Dim tmpMonotributoSoc As Boolean = False
                    Dim tmp As String = lns(i).Substring(45, 2).Trim
                    If tmp.StartsWith("A") Or
                       tmp.StartsWith("B") Or
                       tmp.StartsWith("C") Or
                       tmp.StartsWith("D") Or
                       tmp.StartsWith("E") Or
                       tmp.StartsWith("F") Or
                       tmp.StartsWith("G") Then
                        tmpMonotributo = True

                        If tmp.Length = 2 Then
                            tmpMonotributoSoc = True
                        End If
                    End If

                    If lns(i).Substring(47, 1).Trim = "S" Then
                        .AddColumnValue(TABLA.SOCIEDAD, True)
                    Else
                        .AddColumnValue(TABLA.SOCIEDAD, False)
                    End If

                    If lns(i).Substring(48, 1).Trim = "S" Then
                        .AddColumnValue(TABLA.EMPLEADOR, True)
                    Else
                        .AddColumnValue(TABLA.EMPLEADOR, False)
                    End If

                    .AddColumnValue(TABLA.ACTIVIDAD, CInt(lns(i).Substring(49, 2).Trim))

                    Dim tmpCondicion As Integer = 1

                    If tmpMonotributoSoc Then
                        tmpCondicion = 13
                    Else
                        If tmpMonotributo And (tmpCondicion <> 2) Then
                            tmpCondicion = 6
                        Else
                            Select Case tmpIva
                                Case 1
                                    tmpCondicion = 1
                                Case 2
                                    tmpCondicion = 4
                                Case 4
                                    tmpCondicion = 15
                            End Select
                        End If
                    End If

                    .AddColumnValue(TABLA.CONDICION, tmpCondicion)

                    If Not .Insert() Then Return False
                End With
                i += 1
            Loop Until i >= totalRecords

            If Not silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Actualizar padron AFIP [ OK ]")
            End If
            Return True

        Else
            Return False
        End If
    End Function
End Class

