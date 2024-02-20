Imports helix
Public Class ProductoSocio
	Public Property sqle As New SQLEngine
	Public Property searchResult As New List(Of ProductoSocio)

	Public Property Id As Long = 0
	Public Property SocioId As Long = 0
	Public Property Tipo As Producto_tipo = Producto_tipo.OTROS
	Public Property Descripcion As String = ""
	Public Property Importe As Decimal = 0
	Public Property Movimiento_cc As Long = 0
	Public Property FechaPago As Integer = 0
	Public Property Deleted As Boolean = False
	Public Property Modificado As Date = Now
	Public Enum Guardar As Integer
		NUEVO = 0
		EDITAR = 1
	End Enum

	Public Enum Producto_tipo As Integer
		OTROS = 0
		PUBLICIDAD = 1
		BOLSIN = 2
		MEDICINA_LABORAL = 3
	End Enum


	Public Structure TABLA
		Const TABLA_NOMBRE As String = "ProductoSocio"
		Const ID As String = TABLA_NOMBRE & "_id"
		Const SOCIO_ID As String = TABLA_NOMBRE & "_socioId"
		Const TIPO As String = TABLA_NOMBRE & "_tipo"
		Const DESCRIPCION As String = TABLA_NOMBRE & "_descripcion"
		Const IMPORTE As String = TABLA_NOMBRE & "_importe"
		Const MOVIMIENTO__CC As String = TABLA_NOMBRE & "_movimiento_cc"
		Const FECHA_PAGO As String = TABLA_NOMBRE & "_fechaPago"
		Const DELETED As String = TABLA_NOMBRE & "_deleted"
		Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
		Const ALL As String = ID & ", " & SOCIO_ID & ", " & TIPO & ", " & DESCRIPCION & ", " & IMPORTE & ", " & MOVIMIENTO__CC & ", " & FECHA_PAGO & ", " & DELETED & ", " & MODIFICADO
	End Structure



	Public Sub New()
	End Sub


	Public Sub New(ByVal iSqle As SQLEngine)
		Me.sqle.RequireCredentials = iSqle.RequireCredentials
		Me.sqle.Username = iSqle.Username
		Me.sqle.Password = iSqle.Password
		Me.sqle.dbType = iSqle.dbType
		Me.sqle.Path = iSqle.Path
		Me.sqle.DatabaseName = iSqle.DatabaseName
		If iSqle.IsStarted Then
			Me.sqle.ColdBoot()
		Else
			Me.sqle.Start()
		End If
	End Sub
	Public Sub InjectSQL(ByVal iSqle As SQLEngine)
		Me.sqle.RequireCredentials = iSqle.RequireCredentials
		Me.sqle.Username = iSqle.Username
		Me.sqle.Password = iSqle.Password
		Me.sqle.dbType = iSqle.dbType
		Me.sqle.Path = iSqle.Path
		Me.sqle.DatabaseName = iSqle.DatabaseName
		If iSqle.IsStarted Then
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
					SocioId = CLng(.GetQueryData(TABLA.SOCIO_ID))
					Tipo = CInt(.GetQueryData(TABLA.TIPO))
					Descripcion = CStr(.GetQueryData(TABLA.DESCRIPCION))
					Importe = CDec(.GetQueryData(TABLA.IMPORTE))
					Movimiento_cc = CLng(.GetQueryData(TABLA.MOVIMIENTO__CC))
					FechaPago = CInt(.GetQueryData(TABLA.FECHA_PAGO))
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
					.AddColumnValue(TABLA.SOCIO_ID, SocioId)
					.AddColumnValue(TABLA.TIPO, Tipo)
					.AddColumnValue(TABLA.DESCRIPCION, Descripcion)
					.AddColumnValue(TABLA.IMPORTE, Importe)
					.AddColumnValue(TABLA.MOVIMIENTO__CC, Movimiento_cc)
					.AddColumnValue(TABLA.FECHA_PAGO, FechaPago)
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
					.AddColumnValue(TABLA.SOCIO_ID, SocioId)
					.AddColumnValue(TABLA.TIPO, Tipo)
					.AddColumnValue(TABLA.DESCRIPCION, Descripcion)
					.AddColumnValue(TABLA.IMPORTE, Importe)
					.AddColumnValue(TABLA.MOVIMIENTO__CC, Movimiento_cc)
					.AddColumnValue(TABLA.FECHA_PAGO, FechaPago)
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
	Private Sub LoadQueryData(ByVal currentSqle As SQLEngine, ByRef obj As ProductoSocio)
		With currentSqle.Query
			obj.Id = CLng(.GetQueryData(TABLA.ID))
			obj.SocioId = CLng(.GetQueryData(TABLA.SOCIO_ID))
			obj.Tipo = CInt(.GetQueryData(TABLA.TIPO))
			obj.Descripcion = CStr(.GetQueryData(TABLA.DESCRIPCION))
			obj.Importe = CDec(.GetQueryData(TABLA.IMPORTE))
			obj.Movimiento_cc = CLng(.GetQueryData(TABLA.MOVIMIENTO__CC))
			obj.FechaPago = CInt(.GetQueryData(TABLA.FECHA_PAGO))
			obj.Deleted = CBool(.GetQueryData(TABLA.DELETED))
			obj.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))
		End With
	End Sub
	Public Function QuickSearch(ByVal columna As String, ByVal operador As SQLEngineQuery.OperatorCriteria, ByVal value As Object) As Integer
		With sqle.Query
			.Reset()
			.TableName = TABLA.TABLA_NOMBRE
			.AddSelectColumn(TABLA.ALL)
			.SimpleSearch(columna.ToString, operador, value)
			If .Query() Then
				Me.searchResult.Clear()
				While .QueryRead
					Dim tmp As New ProductoSocio
					LoadQueryData(sqle, tmp)
					searchResult.Add(tmp)
				End While
				Return .RecordCount
			End If
		End With
		Return 0
	End Function

	Public Function LoadPorMovimiento(ByVal movimientoID As Integer) As List(Of ProductoSocio)
		Dim res As New List(Of ProductoSocio)
		With sqle.Query
			.Reset()
			.TableName = TABLA.TABLA_NOMBRE
			.AddSelectColumn(TABLA.ID)
			.WHEREstring = $"{TABLA.MOVIMIENTO__CC} = { .p(movimientoID)} AND 
							 {TABLA.DELETED} = { .p(False)}"
			If .Query Then
				While .QueryRead
					Dim p As New ProductoSocio(Me.sqle)
					p.LoadMe(.GetQueryData(TABLA.ID))
					res.Add(p)
				End While
			End If
		End With

		Return res
	End Function
End Class

