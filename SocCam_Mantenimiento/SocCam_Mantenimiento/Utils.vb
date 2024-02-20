Imports System.Drawing
Imports System.IO
Imports System.Security.Cryptography




Public Class Utils

    Public Shared Function DescargarArchivo(ByVal url As String, ByVal path As String, Optional silent As Boolean = False) As Boolean
        Dim dir As New Uri(url)
        Try
            My.Computer.Network.DownloadFile(dir, $"{path}\{url.Substring(url.LastIndexOf("/"), (url.Length - url.LastIndexOf("/")))}", Nothing, Not silent, 500, True)
            Return True
        Catch ex As Exception
            Debug.Print(ex.Message)
            Return False
        End Try
    End Function

    Public Shared Function LeftPad(ByVal str As String, ByVal tamanio As Integer, ByVal caracter As String) As String
        Dim output As String = str
        For i = str.Length To (tamanio - 1)
            output = caracter & output
        Next
        Return output
    End Function

    Public Shared Function RightPad(ByVal str As String, ByVal tamanio As Integer, ByVal caracter As String) As String
        Dim output As String = str
        For i = str.Length To (tamanio - 1)
            output = output & caracter
        Next
        Return output
    End Function

    Public Shared Function DateTo8601(ByVal fecha As Date) As Integer
        Return CInt(fecha.Year & LeftPad(fecha.Month.ToString, 2, "0") & LeftPad(fecha.Day.ToString, 2, "0"))
    End Function

    Public Shared Function Int8601ToDate(ByVal fecha As Integer) As Date
        If (fecha > 20000000) Then
            Return Date.Parse(fecha.ToString.Insert(6, "-").Insert(4, "-"))
        Else
            Return Now
        End If
    End Function

    Public Shared Function TimeTo8601(ByVal hora As Date) As Integer
        Return CInt(LeftPad(hora.Hour, 2, "0") & LeftPad(hora.Minute, 2, "0") & LeftPad(hora.Second, 2, "0"))
    End Function

    Public Shared Function ComponerNumeroComprobante(ByVal pventa As Integer, ByVal numero As Integer) As String
        Return LeftPad(pventa, 4, "0") & "-" & LeftPad(numero, 8, "0")
    End Function

    Public Shared Function IvaToIndex(ByVal condicion As String) As Integer
        Select Case condicion
            Case "N/A"
                Return 0
            Case "Monotributo"
                Return 1
            Case "Responsable Inscripto"
                Return 2
            Case "Exento"
                Return 3
            Case "Consumidor Final"
                Return 4
            Case "No alcanzado"
                Return 5
            Case Else
                Return 0
        End Select
    End Function


    Public Shared Function ToMoneyFormat(ByVal val As Decimal) As String
        Dim out As String = String.Format("{0:C2}", val)

        If Not out.Contains("$") Then
            Dim p1 As String = out.Split(" ")(0).Trim.Trim.Replace(",", ".")
            Dim p2 As String = out.Split(" ")(1).Trim.Replace(",", ".")
            If IsNumeric(p1) Then
                out = $"$ {p1.Remove(p1.Length - 3, 1).Insert(p1.Length - 3, ",")}"
            Else
                out = $"$ {p2.Remove(p2.Length - 3, 1).Insert(p2.Length - 3, ",")}"
            End If
            If val < 0 Then
                'out = out.Replace("-", "").Replace("$ ", "$ -")
            End If
        End If
        Return out
    End Function

    Public Shared Function ToPercenteFormat(ByVal val As Decimal) As String
        val = val / 100
        Return val.ToString("P")
    End Function


    Public Shared Function GetMesesEntreFechas(ByVal Desde As Date, ByVal Hasta As Date) As Integer
        Return DateDiff(DateInterval.Month, Desde, Hasta)
    End Function

    ''' <summary>
    ''' Calcula el primer día del mes
    ''' </summary>
    ''' <param name="mes">Mes en base 0</param>
    ''' <param name="anio">Anio a calcular</param>
    ''' <returns></returns>
    Public Shared Function GetPrimerDiaMes(ByVal mes As Integer, ByVal anio As Integer) As Date
        Dim tmpMes As String = ""

        If mes < 9 Then
            tmpMes = "0" & (mes + 1).ToString
        Else
            tmpMes = (mes + 1).ToString
        End If

        Dim tmpDate As Date
        Date.TryParse("01/" & tmpMes & "/" & anio, tmpDate)
        Return tmpDate.Date
    End Function

    ''' <summary>
    ''' Calcula el primer día del mes
    ''' </summary>
    ''' <param name="mes">Mes en base 0</param>
    ''' <param name="anio">Anio a calcular</param>
    ''' <returns></returns>
    Public Shared Function GetPrimerDiaMesISO(ByVal mes As Integer, ByVal anio As Integer) As Integer
        Dim tmpMes As String = ""

        If mes < 9 Then
            tmpMes = "0" & (mes + 1).ToString
        Else
            tmpMes = (mes + 1).ToString
        End If


        Return Integer.Parse(anio & tmpMes & "01")
    End Function

    ''' <summary>
    ''' Obtener el ultimo dia del mes
    ''' </summary>
    ''' <param name="mes">Mes del año en base 0</param>
    ''' <param name="anio">Año a obtener</param>
    ''' <returns>El ultimo dia del mes del año especificado</returns>
    Public Shared Function GetUltimoDiaMes(ByVal mes As Integer, ByVal anio As Integer) As Date

        '0 enero, con 31 días
        '1 febrero, con 28 días o 29 en año bisiesto
        '2 marzo, con 31 días
        '3 abril, con 30 días
        '4 mayo, con 31 días
        '5 junio, con 30 días
        '6 julio, con 31 días
        '7 agosto, con 31 días
        '8 septiembre, con 30 días
        '9 octubre, con 31 días
        '10 noviembre, con 30 días
        '11 diciembre, con 31 días

        Dim tmpDiasMes As Integer = 0
        Dim tmpMes As String = ""

        If mes < 9 Then
            tmpMes = "0" & (mes + 1).ToString
        Else
            tmpMes = (mes + 1).ToString
        End If

        If (mes = 0) Or (mes = 2) Or (mes = 4) Or (mes = 6) Or (mes = 7) Or (mes = 9) Or (mes = 11) Then tmpDiasMes = 31
        If (mes = 3) Or (mes = 5) Or (mes = 8) Or (mes = 10) Then tmpDiasMes = 30
        If (mes = 1) Then
            If Date.IsLeapYear(anio) Then
                tmpDiasMes = 29
            Else
                tmpDiasMes = 28
            End If
        End If

        'Dim tmpDate As Date = Date.ParseExact(tmpDiasMes.ToString & "/" & tmpMes & "/" & anio, "dd/mm/yyyy", Globalization.CultureInfo.InvariantCulture)
        'Dim tmpDate As Date
        'Date.TryParse($"{anio}{tmpMes}{tmpDiasMes}", tmpDate)

        Return Utils.Int8601ToDate(CInt($"{anio}{tmpMes}{tmpDiasMes}"))
    End Function

    ''' <summary>
    ''' Obtener el ultimo dia del mes
    ''' </summary>
    ''' <param name="mes">Mes del año en base 0</param>
    ''' <param name="anio">Año a obtener</param>
    ''' <returns>El ultimo dia del mes del año especificado</returns>
    Public Shared Function GetUltimoDiaMesISO(ByVal mes As Integer, ByVal anio As Integer) As Integer

        '0 enero, con 31 días
        '1 febrero, con 28 días o 29 en año bisiesto
        '2 marzo, con 31 días
        '3 abril, con 30 días
        '4 mayo, con 31 días
        '5 junio, con 30 días
        '6 julio, con 31 días
        '7 agosto, con 31 días
        '8 septiembre, con 30 días
        '9 octubre, con 31 días
        '10 noviembre, con 30 días
        '11 diciembre, con 31 días

        Dim tmpDiasMes As Integer = 0
        Dim tmpMes As String = ""

        If mes < 9 Then
            tmpMes = "0" & (mes + 1).ToString
        Else
            tmpMes = (mes + 1).ToString
        End If

        If (mes = 0) Or (mes = 2) Or (mes = 4) Or (mes = 6) Or (mes = 7) Or (mes = 9) Or (mes = 11) Then tmpDiasMes = 31
        If (mes = 3) Or (mes = 5) Or (mes = 8) Or (mes = 10) Then tmpDiasMes = 30
        If (mes = 1) Then
            If Date.IsLeapYear(anio) Then
                tmpDiasMes = 29
            Else
                tmpDiasMes = 28
            End If
        End If

        'Dim tmpDate As Date = Date.ParseExact(tmpDiasMes.ToString & "/" & tmpMes & "/" & anio, "dd/mm/yyyy", Globalization.CultureInfo.InvariantCulture)
        'Dim tmpDate As Date
        'Date.TryParse($"{anio}{tmpMes}{tmpDiasMes}", tmpDate)

        Return Integer.Parse($"{anio}{tmpMes}{tmpDiasMes}")
    End Function


    Public Shared Function ValidarCUIT(ByVal cuit As String) As Boolean
        If IsNumeric(cuit) Then
            Try
                If (cuit.Length <> 11) Then Return False

                Dim tmpCuit As String = cuit.Substring(0, 10)
                Dim total As Integer = 0
                Dim tabla() As Integer = {5, 4, 3, 2, 7, 6, 5, 4, 3, 2}

                For i = 0 To tmpCuit.Length - 1
                    Dim currChar As Integer = CInt(tmpCuit.Substring(i, 1))
                    total += currChar * tabla(i)
                Next
                Dim preDV As Integer = 11 - (total Mod 11)
                If preDV >= 10 Then
                    preDV = 0
                End If
                tmpCuit &= preDV
                If cuit = tmpCuit Then
                    Return True
                Else
                    Return False
                End If

            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Public Shared Function EsDNI(ByVal dni As String) As Boolean
        If Not IsNumeric(dni) Then
            Return False
        Else
            Dim tmpDNI As Long = CLng(dni)
            If tmpDNI > 1000000 And tmpDNI < 50000000 Then
                Return True
            Else
                Return False
            End If
        End If
    End Function


    Public Shared Function AFormatoCUIT(ByVal cuit As String) As String
        If cuit.Length = 11 Then
            Return $"{cuit.Substring(0, 2)}-{cuit.Substring(2, 8)}-{cuit.Substring(10)}"
        Else
            Return cuit
        End If
    End Function

    Public Shared Function CalcularPorcentaje(ByVal total As Decimal, ByVal porcentaje As Decimal) As Decimal
        Return ((total * porcentaje) / 100)
    End Function

    Public Shared Function CalcularMasPorcentaje(ByVal total As Decimal, ByVal porcentaje As Decimal) As Decimal
        Return ((total * porcentaje) / 100) + total
    End Function

    Public Shared Function CalcularMenosPorcentaje(ByVal total As Decimal, ByVal porcentaje As Decimal) As Decimal
        Return total - ((total * porcentaje) / 100)
    End Function

    Public Shared Function ToUTF8(ByVal cadenaOriginal As String) As String
        ' &#195;&#177;
        Dim tmpString As String = System.Net.WebUtility.HtmlEncode(cadenaOriginal)
        tmpString = tmpString.Replace("&#195;&#177;", "ñ").Replace("&#195;&#169;", "é").Replace("&#195;&#173;", "í").Replace("&#195;&#179;", "ó").Replace("&#195;&#188", "u")
        Return tmpString
    End Function

    Public Shared Function ComponerIso8601(ByVal dia As Integer, ByVal mes As Integer, ByVal anio As Integer) As Integer
        Return $"{anio.ToString.PadLeft(4, "0")}{mes.ToString.PadLeft(2, "0")}{dia.ToString.PadLeft(2, "0")}"
    End Function


    Public Shared Function HashMD5(ByVal file_name As String) As String
        Dim hash = MD5.Create()
        Dim hashValue() As Byte
        Dim fileStream As FileStream = File.OpenRead(file_name)
        fileStream.Position = 0
        hashValue = hash.ComputeHash(fileStream)
        Dim hex_value As String = ""
        Dim i As Integer
        For i = 0 To hashValue.Length - 1
            hex_value += hashValue(i).ToString("X2")
        Next i
        fileStream.Close()
        Return hex_value
    End Function




    Public Shared Function NombreProvincia(ByVal idProvincia) As String
        Select Case idProvincia
            Case 0
                Return "Ciudad Autónoma de Buenos Aires"
            Case 1
                Return "Buenos Aires"
            Case 2
                Return "Catamara"
            Case 3
                Return "Córdoba"
            Case 4
                Return "Corrientes"
            Case 5
                Return "Entre Ríos"
            Case 6
                Return "Jujuy"
            Case 7
                Return "Mendoza"
            Case 8
                Return "La Rioja"
            Case 9
                Return "Salta"
            Case 10
                Return "San Juan"
            Case 11
                Return "San Luis"
            Case 12
                Return "Santa Fe"
            Case 13
                Return "Santiago del Estero"
            Case 14
                Return "Tucumán"
            Case 15
                Return "Chaco"
            Case 16
                Return "Chubut"
            Case 17
                Return "Formosa"
            Case 18
                Return "Misiones"
            Case 19
                Return "Neuquén"
            Case 20
                Return "La Pampa"
            Case 21
                Return "Río Negro"
            Case 22
                Return "Santa Cruz"
            Case 23
                Return "Tierra del Fuego"
            Case Else
                Return ""
        End Select

    End Function

    Public Shared Function GetNombreMes(ByVal mes As Integer) As String
        Select Case mes
            Case 0
                Return "Enero"
            Case 1
                Return "Febrero"
            Case 2
                Return "Marzo"
            Case 3
                Return "Abril"
            Case 4
                Return "Mayo"
            Case 5
                Return "Junio"
            Case 6
                Return "Julio"
            Case 7
                Return "Agosto"
            Case 8
                Return "Septiembre"
            Case 9
                Return "Octubre"
            Case 10
                Return "Noviembre"
            Case 11
                Return "Diciembre"
            Case Else
                Return ""
        End Select
    End Function

    Public Shared Function GetNombreBanco(ByVal id As Integer) As String
        Select Case id
            Case 0
                Return "BANCO DE GALICIA Y BUENOS AIRES S.A.U."
            Case 1
                Return "BANCO DE LA NACION ARGENTINA"
            Case 2
                Return "BANCO DE LA PROVINCIA DE BUENOS AIRES"
            Case 3
                Return "INDUSTRIAL And COMMERCIAL BANK OF CHINA"
            Case 4
                Return "CITIBANK N.A."
            Case 5
                Return "BANCO BBVA ARGENTINA S.A."
            Case 6
                Return "BANCO DE LA PROVINCIA DE CORDOBA S.A."
            Case 7
                Return "BANCO SUPERVIELLE S.A."
            Case 8
                Return "BANCO DE LA CIUDAD DE BUENOS AIRES"
            Case 9
                Return "BANCO PATAGONIA S.A."
            Case 10
                Return "BANCO HIPOTECARIO S.A."
            Case 11
                Return "BANCO DE SAN JUAN S.A."
            Case 12
                Return "BANCO MUNICIPAL DE ROSARIO"
            Case 13
                Return "BANCO SANTANDER RIO S.A."
            Case 14
                Return "BANCO DEL CHUBUT S.A."
            Case 15
                Return "BANCO DE SANTA CRUZ S.A."
            Case 16
                Return "BANCO DE LA PAMPA SOCIEDAD DE ECONOMÍA M"
            Case 17
                Return "BANCO DE CORRIENTES S.A."
            Case 18
                Return "BANCO PROVINCIA DEL NEUQUÉN SOCIEDAD ANÓ"
            Case 19
                Return "BRUBANK S.A.U."
            Case 20
                Return "BANCO INTERFINANZAS S.A."
            Case 21
                Return "HSBC BANK ARGENTINA S.A."
            Case 22
                Return "JPMORGAN CHASE BANK, NATIONAL ASSOCIATIO"
            Case 23
                Return "BANCO CREDICOOP COOPERATIVO LIMITADO"
            Case 24
                Return "BANCO DE VALORES S.A."
            Case 25
                Return "BANCO ROELA S.A."
            Case 26
                Return "BANCO MARIVA S.A."
            Case 27
                Return "BANCO ITAU ARGENTINA S.A."
            Case 28
                Return "BANK OF AMERICA, NATIONAL ASSOCIATION"
            Case 29
                Return "BNP PARIBAS"
            Case 30
                Return "BANCO PROVINCIA DE TIERRA DEL FUEGO"
            Case 31
                Return "BANCO DE LA REPUBLICA ORIENTAL DEL URUGU"
            Case 32
                Return "BANCO SAENZ S.A."
            Case 33
                Return "BANCO MERIDIAN S.A."
            Case 34
                Return "BANCO MACRO S.A."
            Case 34
                Return "BANCO COMAFI SOCIEDAD ANONIMA"
            Case 36
                Return "BANCO DE INVERSION Y COMERCIO EXTERIOR S"
            Case 37
                Return "BANCO PIANO S.A."
            Case 38
                Return "BANCO JULIO SOCIEDAD ANONIMA"
            Case 39
                Return "BANCO RIOJA SOCIEDAD ANONIMA UNIPERSONAL"
            Case 40
                Return "BANCO DEL SOL S.A."
            Case 41
                Return "NUEVO BANCO DEL CHACO S. A."
            Case 42
                Return "BANCO VOII S.A."
            Case 43
                Return "BANCO DE FORMOSA S.A."
            Case 44
                Return "BANCO CMF S.A."
            Case 45
                Return "BANCO DE SANTIAGO DEL ESTERO S.A."
            Case 46
                Return "BANCO INDUSTRIAL S.A."
            Case 47
                Return "NUEVO BANCO DE SANTA FE SOCIEDAD ANONIMA"
            Case 48
                Return "BANCO CETELEM ARGENTINA S.A."
            Case 49
                Return "BANCO DE SERVICIOS FINANCIEROS S.A."
            Case 50
                Return "BANCO BRADESCO ARGENTINA S.A.U."
            Case 51
                Return "BANCO DE SERVICIOS Y TRANSACCIONES S.A."
            Case 52
                Return "RCI BANQUE S.A."
            Case 53
                Return "BACS BANCO DE CREDITO Y SECURITIZACION S"
            Case 54
                Return "BANCO MASVENTAS S.A."
            Case 55
                Return "WILOBANK S.A."
            Case 56
                Return "NUEVO BANCO DE ENTRE RÍOS S.A."
            Case 57
                Return "BANCO COLUMBIA S.A."
            Case 58
                Return "BANCO BICA S.A."
            Case 59
                Return "BANCO COINAG S.A."
            Case 60
                Return "BANCO DE COMERCIO S.A."
            Case 61
                Return "BANCO SUCREDITO REGIONAL S.A.U."
            Case 62
                Return "FORD CREDIT COMPAÑIA FINANCIERA S.A."
            Case 63
                Return "COMPAÑIA FINANCIERA ARGENTINA S.A."
            Case 64
                Return "VOLKSWAGEN Financial SERVICES COMPAÑIA F"
            Case 65
                Return "CORDIAL COMPAÑÍA FINANCIERA S.A."
            Case 66
                Return "FCA COMPAÑIA FINANCIERA S.A."
            Case 67
                Return "GPAT COMPAÑIA FINANCIERA S.A.U."
            Case 68
                Return "MERCEDES-BENZ COMPAÑÍA FINANCIERA ARGENT"
            Case 69
                Return "ROMBO COMPAÑÍA FINANCIERA S.A."
            Case 70
                Return "JOHN DEERE CREDIT COMPAÑÍA FINANCIERA S."
            Case 71
                Return "PSA FINANCE ARGENTINA COMPAÑÍA FINANCIER"
            Case 72
                Return "TOYOTA COMPAÑÍA FINANCIERA DE ARGENTINA"
            Case 73
                Return "FINANDINO COMPAÑIA FINANCIERA S.A."
            Case 74
                Return "MONTEMAR COMPAÑIA FINANCIERA S.A."
            Case 75
                Return "TRANSATLANTICA COMPAÑIA FINANCIERA S.A."
            Case 76
                Return "CREDITO REGIONAL COMPAÑIA FINANCIERA S.A"
            Case 77
                Return "N/A"
            Case Else
                Return "N/A"
        End Select
    End Function


    ''' <summary>
    ''' Convierte un array de Bytes en una imagen
    ''' </summary>
    ''' <param name="avatar">Imagen donde se va a dibujar el array de Bytes</param>
    ''' <param name="dbImage">Array de Bytes con la informacion de una imagen</param>
    ''' <remarks></remarks>
    Public Shared Sub ByteToImage(ByRef avatar As Image, ByVal dbImage As Byte())
        If Not IsDBNull(avatar) Then
            Dim memoryStream As New MemoryStream(dbImage)
            avatar = Image.FromStream(memoryStream)
        End If
    End Sub

    ''' <summary>
    ''' Convierte la imagen de un PictureBox a un array de Bytes
    ''' </summary>
    ''' <param name="avatar">Imagen a ser convertida</param>
    ''' <returns>Un array de Bytes con los datos de una imagen</returns>
    ''' <remarks>Si el PictureBox tiene recuadro, este es copiado. Lo ideal es llamar
    '''  a la funcion con el borde deshabilitado</remarks>
    Public Shared Function ImageToByte(ByVal avatar As Image) As Byte()
        If Not IsNothing(avatar) Then
            Dim memoryStream As New MemoryStream
            avatar.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png)
            Return memoryStream.ToArray()
        Else
            Return Nothing
        End If
    End Function

    Public Shared Sub Scream(ByVal message As String, Optional ByVal errorType As MsgBoxStyle = MsgBoxStyle.Exclamation)
        MsgBox(message, MsgBoxStyle.OkOnly + errorType, "Error")
    End Sub

    ''' <summary>
    ''' Muestra el periodo en base 0 segun la fecha del año y la periodicidad
    ''' </summary>
    ''' <param name="fecha">Fecha a verificar el periodo</param>
    ''' <param name="mesesPorPeriodo">Cantidad de meses que tiene el periodo</param>
    ''' <returns>El periodo que esta incluida la fecha, si la fecha es 1/4/2016 y el periodo es semestral retorna 0</returns>
    ''' <remarks></remarks>
    Public Shared Function GetPeriodoFromFecha(ByVal fecha As Date, ByVal mesesPorPeriodo As Byte) As Integer
        Return ((fecha.Month - 1) \ mesesPorPeriodo)
    End Function


    Public Shared Function GetPeriodoFromFecha(ByVal mes As Integer, ByVal mesesPorPeriodo As Byte) As Integer
        Return ((mes) \ mesesPorPeriodo)
    End Function

    ''' <summary>
    ''' Calcula la cantidad de meses que tiene un peridodo
    ''' </summary>
    ''' <param name="periodicidad">La periodicidad a calcular</param>
    ''' <returns>La cantidad de meses que tiene un periodo</returns>
    Public Shared Function GetMesesFromPeriodicidad(ByVal periodicidad) As Integer
        Select Case periodicidad
            Case 0
                Return 1
            Case 1
                Return 2
            Case 2
                Return 3
            Case 3
                Return 4
            Case 4
                Return 6
            Case 5
                Return 12
        End Select
    End Function




    Public Shared Function TraducirCondicionFiscal(Optional comboIndex As Integer = -1, Optional dbIndex As Integer = -1) As Integer
        If comboIndex <> -1 Then
            Select Case comboIndex
                Case 0
                    Return 1
                Case 1
                    Return 4
                Case 2
                    Return 5
                Case 3
                    Return 6
                Case 4
                    Return 8
                Case 5
                    Return 9
                Case 6
                    Return 10
                Case 7
                    Return 11
                Case 8
                    Return 13
                Case 9
                    Return 15
                Case Else
                    Return -1
            End Select
        End If
        If dbIndex <> -1 Then
            Select Case dbIndex
                Case 1
                    Return 0
                Case 4
                    Return 1
                Case 5
                    Return 2
                Case 6
                    Return 3
                Case 8
                    Return 4
                Case 9
                    Return 5
                Case 10
                    Return 6
                Case 11
                    Return 7
                Case 13
                    Return 8
                Case 15
                    Return 9
                Case Else
                    Return 15
            End Select
        End If

        ' Si llega hasta aca estamos al horno
        Return -1
    End Function

    Public Shared Function GetCondicionFiscalString(ByVal condicionFiscal As Integer) As String
        Select Case condicionFiscal
            Case 1
                Return "IVA Responsable Inscripto"
            Case 4
                Return "IVA Sujeto Exento"
            Case 5
                Return "Consumidor Final"
            Case 6
                Return "Responsable Monotributo"
            Case 8
                Return "Proveedor del Exterior"
            Case 9
                Return "Cliente del Exterior"
            Case 10
                Return "IVA Liberado - Ley Nº 19.640"
            Case 11
                Return "IVA Responsable Inscripto - Agente de Percepción"
            Case 13
                Return "Monotributista Social"
            Case Else
                Return "IVA No Alcanzado"

        End Select
    End Function

    Public Shared Function ToSentenceCase(ByVal str As String) As String
        Dim tmpStr As String = str.ToLower.Trim
        If tmpStr.Length > 0 Then
            tmpStr = tmpStr.Substring(0, 1).ToUpper & tmpStr.Substring(1, tmpStr.Length - 1).Replace("  ", " ")
            Dim space As Boolean = False
            Dim count As Integer = 0
            For Each letra In tmpStr
                If letra = " " Then
                    space = True
                Else
                    If space = True Then
                        tmpStr = tmpStr.Insert(count, tmpStr.Substring(count, 1).ToUpper)
                        tmpStr = tmpStr.Remove(count + 1, 1)
                        space = False
                    End If
                End If
                count += 1
            Next
        End If

        Return tmpStr
    End Function



    ''' <summary>
    ''' Calcula el prorcional de un importe entre dos fechas, segun una fecha dentro del rango
    ''' </summary>
    ''' <param name="fechaDesde"></param>
    ''' <param name="fechaHasta"></param>
    ''' <param name="fechaCalculo"></param>
    ''' <param name="importe"></param>
    ''' <returns></returns>
    Public Shared Function CalcularProporcional(ByVal fechaDesde As Date, ByVal fechaHasta As Date, ByVal fechaCalculo As Date, ByVal importe As Decimal) As Decimal
        Dim totalDias As Integer = DateDiff(DateInterval.Day, fechaDesde, fechaHasta)
        Dim faltanDias As Integer = DateDiff(DateInterval.Day, fechaCalculo, fechaHasta)

        If fechaCalculo < fechaDesde Then Return importe
        If fechaCalculo > fechaHasta Then Return 0

        Dim porcentajeDias As Decimal = (faltanDias * 100) / totalDias

        Return Utils.CalcularPorcentaje(importe, porcentajeDias)
    End Function

    Public Shared Function LimpiarDetalleFactura(ByVal detalle As AfipFacturaDetalle) As AfipFacturaDetalle
        ' Limpiar cuota sociales
        Dim det As New AfipFacturaDetalle()
        det = detalle
        Dim prodServ As String = detalle.ProductoServicio

        Select Case detalle.Codigo
            Case 0
                det.Cantidad = CInt(prodServ.Substring(0, prodServ.IndexOf(" ")))
                det.ProductoServicio = If(det.Cantidad > 1, "Cuotas Sociales", "Cuota Social")
                Dim tmp As String = prodServ.Substring(prodServ.IndexOf("$")).Replace("</b>", "").Replace(".", "").Replace("$", "").Trim
                If det.Cantidad = 0 Then det.Cantidad = 1
                det.PrecioUnitario = CDec(tmp) / det.Cantidad
            Case 1
                det.Cantidad = 1
                det.ProductoServicio = prodServ.Substring(0, prodServ.IndexOf(":"))
                Dim tmp As String = prodServ.Substring(prodServ.IndexOf("$")).Replace("</b>", "").Replace(".", "").Replace("$", "").Trim
                det.PrecioUnitario = CDec(tmp) / det.Cantidad
            Case Else
                det.Cantidad = 1
                det.ProductoServicio = prodServ.Substring(0, prodServ.IndexOf(":"))
                Dim tmp As String = prodServ.Substring(prodServ.IndexOf("$")).Replace("</b>", "").Replace(".", "").Replace("$", "").Trim
                If det.Cantidad = 0 Then det.Cantidad = 1
                det.PrecioUnitario = CDec(tmp) / det.Cantidad
        End Select

        Return det
    End Function
End Class
