Imports System.Text

Public Class AfipQREncoder

    Public Const BASE_URL As String = "https://www.afip.gob.ar/fe/qr/?p="

    ''' <summary>
    ''' OBLIGATORIO – versión del formato de los datos del comprobante
    ''' </summary>
    Public Const VER As String = "1"

    ''' <summary>
    ''' Tipo de código de autorización CAEA 
    ''' </summary>
    Public Const TIPO_CAEA As String = "A"
    ''' <summary>
    ''' Tipo de código de autorización CAE
    ''' </summary>
    Public Const TIPO_CAE As String = "E"



    ''' <summary>
    ''' OBLIGATORIO – Fecha de emisión del comprobante
    ''' </summary>
    ''' <returns></returns>
    Public Property Fecha As Date

    ''' <summary>
    ''' OBLIGATORIO – Cuit del Emisor del comprobante
    ''' </summary>
    ''' <returns></returns>
    Public Property Cuit As Long

    ''' <summary>
    ''' OBLIGATORIO – Punto de venta utilizado para emitir el comprobante
    ''' </summary>
    ''' <returns></returns>
    Public Property PtoVta As Integer

    ''' <summary>
    ''' OBLIGATORIO – tipo de comprobante
    ''' </summary>
    ''' <see cref="https://www.afip.gob.ar/fe/ayuda/tablas.asp"/>
    ''' <returns></returns>
    Public Property TipoCmp As Integer

    ''' <summary>
    ''' OBLIGATORIO – Número del comprobante
    ''' </summary>
    ''' <returns></returns>
    Public Property NroCmp As Long

    ''' <summary>
    ''' OBLIGATORIO – Importe Total del comprobante (en la moneda en la que fue emitido)
    ''' </summary>
    ''' <returns></returns>
    Public Property Importe As Decimal

    ''' <summary>
    ''' OBLIGATORIO – Moneda del comprobante
    ''' </summary>
    ''' <see cref="https://www.afip.gob.ar/fe/ayuda/tablas.asp"/>
    ''' <returns></returns>
    Public Property Moneda As String

    ''' <summary>
    ''' OBLIGATORIO – Cotización en pesos argentinos de la moneda utilizada (1 cuando la moneda sea pesos)
    ''' </summary>
    ''' <returns></returns>
    Public Property Ctz As Decimal

    ''' <summary>
    ''' DE CORRESPONDER – Código del Tipo de documento del receptor
    ''' </summary>
    ''' <returns></returns>
    Public Property TipoDocRec As Integer

    ''' <summary>
    ''' DE CORRESPONDER – Número de documento del receptor correspondiente al tipo de documento indicado
    ''' </summary>
    ''' <returns></returns>
    Public Property NroDocRec As Long

    ''' <summary>
    ''' OBLIGATORIO – “A” para comprobante autorizado por CAEA, “E” para comprobante autorizado por CAE
    ''' </summary>
    ''' <returns></returns>
    Public Property TipoCodAut As String

    ''' <summary>
    ''' OBLIGATORIO – Código de autorización otorgado por AFIP para el comprobante
    ''' </summary>
    ''' <returns></returns>
    Public Property CodAut As Long

    ''' <summary>
    ''' Genera un JSON con los datos del objeto
    ''' </summary>
    ''' <returns>Cadena vacía si falló. JSON si lo procesó correctamente</returns>
    Public Function GenerarJSON() As String
        Dim output As New StringBuilder
        output.Append("{")
        output.Append($"""ver"":{VER},")
        output.Append($"""fecha"":""{Fecha.Year}-{Fecha.Month.ToString.PadLeft(2, "0")}-{Fecha.Day.ToString.PadLeft(2, "0")}"",")
        output.Append($"""cuit"":{Cuit},")
        output.Append($"""ptoVta"":{PtoVta},")
        output.Append($"""tipoCmp"":{TipoCmp},")
        output.Append($"""nroCmp"":{NroCmp},")
        output.Append($"""importe"":{Importe.ToString.Replace(",", ".")},") ' AFIP no lo dice pero si el separador decimal es la coma, rompe el JSON
        output.Append($"""moneda"":""{Moneda}"",")
        output.Append($"""ctz"":{Ctz.ToString.Replace(",", ".")},")
        output.Append(If(TipoDocRec, $"""tipoDocRec"":{TipoDocRec},", ""))
        output.Append(If(NroDocRec, $"""nroDocRec"":{NroDocRec},", ""))
        output.Append($"""tipoCodAut"":""{TipoCodAut}"",")
        output.Append($"""codAut"":{CodAut}")
        output.Append("}")

        Return If(output.Length >= 47, output.ToString, "")
    End Function

    ''' <summary>
    ''' Devuelve la URL con el JSON codificado en Base64
    ''' </summary>
    ''' <returns>Cadena vacia si falló, la URL si se completó correctamente</returns>
    Public Function GenerarURLEncoded() As String
        Dim json As String = GenerarJSON()
        If json.Length > 0 Then
            Return BASE_URL + Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes(json))
        End If

        Return ""
    End Function
End Class
