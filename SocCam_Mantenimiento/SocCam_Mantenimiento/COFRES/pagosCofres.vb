Imports helix
Public Class pagosCofre
	Public Property sqle As New SQLEngine
	Public Property searchResult As New List(Of pagosCofre)

	Public Property Id As Long = 0
	Public Property Contrato As Long = 0
	Public Property Importe As Decimal = 0
	Public Property Periodo As Integer = 0
	Public Property Anio As Integer = 0
	Public Property FechaPago As Date = #1/1/1970#
	Public Property Estado As Integer = 0
	Public Property Deleted As Boolean = False
	Public Property Modificado As Date = Now
	Public Enum Guardar
		NUEVO = 0
		EDITAR = 1
	End Enum


	Public Structure TABLA
		Const TABLA_NOMBRE As String = "pagosCofres"
		Const ID As String = TABLA_NOMBRE & "_id"
		Const CONTRATO As String = TABLA_NOMBRE & "_contrato"
		Const IMPORTE As String = TABLA_NOMBRE & "_importe"
		Const PERIODO As String = TABLA_NOMBRE & "_periodo"
		Const ANIO As String = TABLA_NOMBRE & "_anio"
		Const FECHA_PAGO As String = TABLA_NOMBRE & "_fechaPago"
		Const ESTADO As String = TABLA_NOMBRE & "_estado"
		Const DELETED As String = TABLA_NOMBRE & "_deleted"
		Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
		Const ALL As String = ID & ", " & CONTRATO & ", " & IMPORTE & ", " & PERIODO & ", " & ANIO & ", " & FECHA_PAGO & ", " & ESTADO & ", " & DELETED & ", " & MODIFICADO
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
					Contrato = CLng(.GetQueryData(TABLA.CONTRATO))
					Importe = CDec(.GetQueryData(TABLA.IMPORTE))
					Periodo = CInt(.GetQueryData(TABLA.PERIODO))
					Anio = CInt(.GetQueryData(TABLA.ANIO))
					FechaPago = CDate(.GetQueryData(TABLA.FECHA_PAGO))
					Estado = CInt(.GetQueryData(TABLA.ESTADO))
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
					.AddColumnValue(TABLA.CONTRATO, Contrato)
					.AddColumnValue(TABLA.IMPORTE, Importe)
					.AddColumnValue(TABLA.PERIODO, Periodo)
					.AddColumnValue(TABLA.ANIO, Anio)
					.AddColumnValue(TABLA.FECHA_PAGO, FechaPago)
					.AddColumnValue(TABLA.ESTADO, Estado)
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
					.AddColumnValue(TABLA.CONTRATO, Contrato)
					.AddColumnValue(TABLA.IMPORTE, Importe)
					.AddColumnValue(TABLA.PERIODO, Periodo)
					.AddColumnValue(TABLA.ANIO, Anio)
					.AddColumnValue(TABLA.FECHA_PAGO, FechaPago)
					.AddColumnValue(TABLA.ESTADO, Estado)
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
					Dim tmp As New pagosCofre
					tmp.Id = CLng(.GetQueryData(TABLA.ID))
					tmp.Contrato = CLng(.GetQueryData(TABLA.CONTRATO))
					tmp.Importe = CDec(.GetQueryData(TABLA.IMPORTE))
					tmp.Periodo = CInt(.GetQueryData(TABLA.PERIODO))
					tmp.Anio = CInt(.GetQueryData(TABLA.ANIO))
					tmp.FechaPago = CDate(.GetQueryData(TABLA.FECHA_PAGO))
					tmp.Estado = CInt(.GetQueryData(TABLA.ESTADO))
					tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
					tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))
					searchResult.Add(tmp)
				End While
				Return .RecordCount
			End If
		End With
		Return 0
	End Function

	Public Function ActualizarEstadoCuotas(ByVal fechaActual As Date) As Integer
		Dim periodo As Integer = 0
		If fechaActual.Month > 6 Then
			periodo = 1
		End If

		' Ejecutar vencimiento de años anteriores
		With sqle.Update
			.Reset()
			.TableName = TABLA.TABLA_NOMBRE
			.AddColumnValue(TABLA.ESTADO, 2)
			.WHEREstring = $"{TABLA.ANIO} < ? AND {TABLA.ESTADO} <> ?"
			.AddWHEREparam(fechaActual.Year)
			.AddWHEREparam(0)
			If Not .Update() Then Return 1  ' 1 Indica fallo en actualización de años anteriores
		End With

		Dim actualizarVencimientos As Boolean = False

		If periodo = 1 Then
			With sqle.Update
				.Reset()
				.TableName = TABLA.TABLA_NOMBRE
				.AddColumnValue(TABLA.ESTADO, 2)
				.WHEREstring = $"{TABLA.ANIO} = ? AND {TABLA.PERIODO} = ? AND {TABLA.ESTADO} <> ?"
				.AddWHEREparam(fechaActual.Year)
				.AddWHEREparam(0)
				.AddWHEREparam(0)
				If Not .Update() Then Return 2  ' 2 Indica fallo en actualización en primer semestre
			End With


			If fechaActual.Day > 20 Or fechaActual.Month > 7 Then
				actualizarVencimientos = True
			End If
		Else
			If fechaActual.Day > 20 Or fechaActual.Month > 1 Then
				actualizarVencimientos = True
			End If
		End If

		If actualizarVencimientos Then
			With sqle.Update
				.Reset()
				.TableName = TABLA.TABLA_NOMBRE
				.AddColumnValue(TABLA.ESTADO, 2)
				.WHEREstring = $"{TABLA.ANIO} = ? AND {TABLA.PERIODO} = ? AND {TABLA.ESTADO} <> ?"
				.AddWHEREparam(fechaActual.Year)
				.AddWHEREparam(periodo)
				.AddWHEREparam(0)
				If Not .Update() Then Return 3  ' 3 Indica fallo en actualización en el periodo actual
			End With
		End If

		Return 0
	End Function

	''' <summary>
	''' Actualiza la base de datos de ProNET en funcion si tienen deudas o no.
	''' </summary>
	''' <param name="pronetPath"></param>
	''' <returns></returns>
	Public Function ActualizarBloqueos(ByVal pronetPath As String) As Integer
		Return 0
	End Function

End Class

