Public Class DetalleCCSocio

    Public Property ListadoCuotasSocialesVirtuales As New Dictionary(Of Integer, CuotaSocio)

	Public Enum TipoDeMovimiento As Integer
		SOCIOS_CUOTA_SOCIAL = 0
		SOCIOS_OTROS = 1
		SOCIOS_PUBLICIDAD = 2
		SOCIOS_BOLSIN = 3
		SOCIOS_MEDICINA = 4
		SOCIOS_CUOTA_SOCIAL_CAMPANIA = 5
	End Enum

	Public Property Tipo As TipoDeMovimiento
	Public Property Importe As Decimal = 0
	Public Property Descripcion As String = 0
	Public Property Campania As New Campania()
	Public Property IdCuota As Integer = 0

	''' <summary>
	''' Busca en el diccionario de cuotas si existe el periodo
	''' </summary>
	''' <param name="periodo">Periodo en formato 202000 (Enero 2020)</param>
	''' <returns>True si se encuentra, False si no</returns>
	Public Function CuotaSocialExiste(ByVal periodo As Integer) As Boolean
		Return ListadoCuotasSocialesVirtuales.ContainsKey(periodo)
	End Function

	Public Function GetUltimaCuotaSocialVirtual() As CuotaSocio
		Return ListadoCuotasSocialesVirtuales.Item(ListadoCuotasSocialesVirtuales.Count - 1)
	End Function

End Class
