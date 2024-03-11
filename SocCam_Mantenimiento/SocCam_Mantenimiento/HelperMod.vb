Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports helix

Module HelperMod

    Dim ConsoleOut As New ConsoleOut

    ' Public Const SQL_PASS As String = "jqW8lU1uPWoV"
    ' Public Const SQL_USER As String = "sa"

    ' usuario soccam_user
    ' pass: Soccam_2015

    Public Const GUARDAR_NUEVO As Integer = 0
    Public Const GUARDAR_EDITAR As Integer = 1

    ''' <summary>
    ''' Rutina que ayuda a ver que valor se va a cargar en el subitem de un ListItem
    ''' </summary>
    ''' <param name="lstIndex">subitem index</param>
    ''' <param name="tableIndex">index del campo de la consulta a la db</param>
    ''' <param name="value">valor devuelto en la columna tableIndex</param>
    ''' <param name="fieldName">Cadena con el nombre de la columna</param>
    ''' <remarks></remarks>
    Public Sub DebugList(ByVal lstIndex As Integer, ByVal tableIndex As Integer, ByVal value As Object, Optional fieldName As String = "")

        Dim optFieldName As String = ""

        If fieldName.Length <> 0 Then
            optFieldName = " | TBL_NAME: " & fieldName
        End If

    End Sub

    ''' <summary>
    ''' Convierte un array de Bytes en una imagen
    ''' </summary>
    ''' <param name="avatar">Imagen donde se va a dibujar el array de Bytes</param>
    ''' <param name="dbImage">Array de Bytes con la informacion de una imagen</param>
    ''' <remarks></remarks>
    Public Sub ByteToImage(ByRef avatar As Image, ByVal dbImage As Byte())
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
    Public Function ImageToByte(ByVal avatar As Image) As Byte()
        If Not IsNothing(avatar) Then
            Dim memoryStream As New MemoryStream
            avatar.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png)
            Return memoryStream.ToArray()
        Else
            Return Nothing
        End If
    End Function


    Public Function ImageToBase64(ByVal avatar As Image) As String
        If Not IsNothing(avatar) Then
            Dim memoryStream As New MemoryStream
            avatar.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png)
            Return Convert.ToBase64String(memoryStream.ToArray())
        Else
            Return ""
        End If

    End Function

    ''' <summary>
    ''' Retorna el nombre del mes según el byte ingresado
    ''' </summary>
    ''' <param name="mes">valor en base 0 con el mes a retornar el nombre</param>
    ''' <returns>El nombre del mes solicitado</returns>
    ''' <remarks>Mes es en base 0, si está fuera de rango devuelve enero</remarks>
    Public Function GetNombreMes(ByVal mes As Byte) As String
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
                Return "Enero"
        End Select
    End Function



    Public Function ToSentenceCase(ByVal str As String) As String
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

    Public Function CamelCaseToUnderscore(ByVal str As String) As String
        Dim count As Integer = 0
        Dim finalString As String = str
        Dim inserts As New List(Of Integer)
        For Each letra In str
            If letra.ToString = str.Substring(count, 1).ToUpper Then
                inserts.Add(count)
            End If
            count += 1
        Next

        If inserts.Count > 0 Then
            inserts.Reverse()
            count = 0
            For i = str.Length - 1 To 0 Step -1
                If i = inserts(count) Then
                    finalString = finalString.Insert(i, "_")
                    count += 1
                    If count >= inserts.Count Then Exit For
                End If
            Next
        End If

        Return finalString
    End Function

    Public Function ToAfipMoneyFormat(ByVal val As Decimal) As String
        Return val.ToString("F")
    End Function

    Public Function ToPercenteFormat(ByVal val As Decimal) As String
        val = val / 100
        Return val.ToString("P")
    End Function

    Public Function ToCuitConGuiones(ByVal val As String) As String
        If val.Length = 11 Then
            Return $"{val.Substring(0, 2)}-{val.Substring(2, 8)}-{val.Substring(10, 1)}"
        Else
            Return val
        End If
    End Function


    '****************************************
    'Desarrollado por: Pedro Alex Taya Yactayo
    'Email: alextaya@hotmail.com
    'Web: http://es.geocities.com/wiseman_alextaya
    '     http://groups.msn.com/mugcanete
    '****************************************

    Public Function NumeroALetras(ByVal numero As String) As String
        '********Declara variables de tipo cadena************
        Dim palabras As String = ""
        Dim entero As String = ""
        Dim dec As String = ""
        Dim flag As String = ""

        '********Declara variables de tipo entero***********
        Dim num, x, y As Integer

        flag = "N"

        '**********Número Negativo***********
        If Mid(numero, 1, 1) = "-" Then
            numero = Mid(numero, 2, numero.ToString.Length - 1).ToString
            palabras = "menos "
        End If

        '**********Si tiene ceros a la izquierda*************
        For x = 1 To numero.ToString.Length
            If Mid(numero, 1, 1) = "0" Then
                numero = Trim(Mid(numero, 2, numero.ToString.Length).ToString)
                If Trim(numero.ToString.Length) = 0 Then palabras = ""
            Else
                Exit For
            End If
        Next

        '*********Dividir parte entera y decimal************
        For y = 1 To Len(numero)
            If (Mid(numero, y, 1) = ".") Or (Mid(numero, y, 1) = ",") Then
                flag = "S"
            Else
                If flag = "N" Then
                    entero = entero + Mid(numero, y, 1)
                Else
                    dec = dec + Mid(numero, y, 1)
                End If
            End If
        Next y

        If Len(dec) = 1 Then dec = dec & "0"

        '**********proceso de conversión***********
        flag = "N"

        If Val(numero) <= 999999999 Then
            For y = Len(entero) To 1 Step -1
                num = Len(entero) - (y - 1)
                Select Case y
                    Case 3, 6, 9
                        '**********Asigna las palabras para las centenas***********
                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If Mid(entero, num + 1, 1) = "0" And Mid(entero, num + 2, 1) = "0" Then
                                    palabras = palabras & "cien "
                                Else
                                    palabras = palabras & "ciento "
                                End If
                            Case "2"
                                palabras = palabras & "doscientos "
                            Case "3"
                                palabras = palabras & "trescientos "
                            Case "4"
                                palabras = palabras & "cuatrocientos "
                            Case "5"
                                palabras = palabras & "quinientos "
                            Case "6"
                                palabras = palabras & "seiscientos "
                            Case "7"
                                palabras = palabras & "setecientos "
                            Case "8"
                                palabras = palabras & "ochocientos "
                            Case "9"
                                palabras = palabras & "novecientos "
                        End Select
                    Case 2, 5, 8
                        '*********Asigna las palabras para las decenas************
                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    flag = "S"
                                    palabras = palabras & "diez "
                                End If
                                If Mid(entero, num + 1, 1) = "1" Then
                                    flag = "S"
                                    palabras = palabras & "once "
                                End If
                                If Mid(entero, num + 1, 1) = "2" Then
                                    flag = "S"
                                    palabras = palabras & "doce "
                                End If
                                If Mid(entero, num + 1, 1) = "3" Then
                                    flag = "S"
                                    palabras = palabras & "trece "
                                End If
                                If Mid(entero, num + 1, 1) = "4" Then
                                    flag = "S"
                                    palabras = palabras & "catorce "
                                End If
                                If Mid(entero, num + 1, 1) = "5" Then
                                    flag = "S"
                                    palabras = palabras & "quince "
                                End If
                                If Mid(entero, num + 1, 1) > "5" Then
                                    flag = "N"
                                    palabras = palabras & "dieci"
                                End If
                            Case "2"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "veinte "
                                    flag = "S"
                                Else
                                    palabras = palabras & "veinti"
                                    flag = "N"
                                End If
                            Case "3"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "treinta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "treinta y "
                                    flag = "N"
                                End If
                            Case "4"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "cuarenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "cuarenta y "
                                    flag = "N"
                                End If
                            Case "5"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "cincuenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "cincuenta y "
                                    flag = "N"
                                End If
                            Case "6"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "sesenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "sesenta y "
                                    flag = "N"
                                End If
                            Case "7"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "setenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "setenta y "
                                    flag = "N"
                                End If
                            Case "8"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "ochenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "ochenta y "
                                    flag = "N"
                                End If
                            Case "9"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "noventa "
                                    flag = "S"
                                Else
                                    palabras = palabras & "noventa y "
                                    flag = "N"
                                End If
                        End Select
                    Case 1, 4, 7
                        '*********Asigna las palabras para las unidades*********
                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If flag = "N" Then
                                    If y = 1 Then
                                        palabras = palabras & "uno "
                                    Else
                                        palabras = palabras & "un "
                                    End If
                                End If
                            Case "2"
                                If flag = "N" Then palabras = palabras & "dos "
                            Case "3"
                                If flag = "N" Then palabras = palabras & "tres "
                            Case "4"
                                If flag = "N" Then palabras = palabras & "cuatro "
                            Case "5"
                                If flag = "N" Then palabras = palabras & "cinco "
                            Case "6"
                                If flag = "N" Then palabras = palabras & "seis "
                            Case "7"
                                If flag = "N" Then palabras = palabras & "siete "
                            Case "8"
                                If flag = "N" Then palabras = palabras & "ocho "
                            Case "9"
                                If flag = "N" Then palabras = palabras & "nueve "
                        End Select
                End Select

                '***********Asigna la palabra mil***************
                If y = 4 Then
                    If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or
                    (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And
                    Len(entero) <= 6) Then palabras = palabras & "mil "
                End If

                '**********Asigna la palabra millón*************
                If y = 7 Then
                    If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                        palabras = palabras & "millón "
                    Else
                        palabras = palabras & "millones "
                    End If
                End If
            Next y

            '**********Une la parte entera y la parte decimal*************
            If dec <> "" And dec <> "0000000000000" Then
                NumeroALetras = palabras & "con " & NumeroALetras(dec)
            Else
                NumeroALetras = palabras
            End If
        Else
            NumeroALetras = ""
        End If
    End Function

    Public Function ToFormNum(ByVal num As Integer) As String
        Dim tmpStr As String = num.ToString
        For i = num.ToString.Length To 7
            tmpStr = "0" & tmpStr
        Next
        Return tmpStr
    End Function


    Public Function FileBackup(ByVal fileOrig As String,
                               ByVal dirBackup As String,
                               ByVal checkCRC As Boolean) As Boolean

        If fileOrig.Length = 0 Or dirBackup.Length = 0 Then Return False

        Dim crcOrig As Integer = 0
        Dim crcBackup As Integer = 0
        Dim fileOrigFilename As String = fileOrig.Substring(fileOrig.LastIndexOf("\"), fileOrig.Length - fileOrig.LastIndexOf("\"))
        Dim fileDest As String = dirBackup & fileOrigFilename & "." & Now.ToString.Replace("/", "_").Replace(":", "").Replace(" ", "").Replace("a.m.", "").Replace("p.m.", "").Trim & ".bak"
        Try


            ' Primero verificar que el archivo de origen exista
            If My.Computer.FileSystem.FileExists(fileOrig) Then

                My.Computer.FileSystem.CopyFile(fileOrig, fileDest)

                If checkCRC Then
                    Dim c As New CRC32()
                    ' CRC32 Hash del origen:
                    Dim f As FileStream = New FileStream(fileOrig, FileMode.Open,
                     FileAccess.Read, FileShare.Read, 8192)
                    crcOrig = c.GetCrc32(f)
                    f.Close()

                    Dim c2 As New CRC32()
                    Dim f2 As FileStream = New FileStream(fileDest, FileMode.Open,
                     FileAccess.Read, FileShare.Read, 8192)
                    crcBackup = c2.GetCrc32(f2)
                    f2.Close()

                    If crcOrig <> crcBackup Then
                        Return False
                    End If

                End If

                Return True

            End If

            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Muestra el periodo en base 0 segun la fecha del año y la periodicidad
    ''' </summary>
    ''' <param name="fecha">Fecha a verificar el periodo</param>
    ''' <param name="mesesPorPeriodo">Cantidad de meses que tiene el periodo</param>
    ''' <returns>El periodo que esta incluida la fecha, si la fecha es 1/4/2016 y el periodo es semestral retorna 0</returns>
    ''' <remarks></remarks>
    Public Function GetPeriodoFromFecha(ByVal fecha As Date, ByVal mesesPorPeriodo As Byte) As Integer
        Return ((fecha.Month - 1) \ mesesPorPeriodo)
    End Function


    Public Function GetPeriodoFromFecha(ByVal mes As Integer, ByVal mesesPorPeriodo As Byte) As Integer
        Return ((mes) \ mesesPorPeriodo)
    End Function



    ''' <summary>
    ''' Convierte una lista de booleanos en un entero
    ''' </summary>
    ''' <param name="lstFlags"></param>
    ''' <returns></returns>
    Public Function FlagsToInt(ByVal lstFlags As List(Of Boolean)) As Integer
        Dim tmpInt As Integer = 0
        Dim counter As Integer = 0
        For Each bool In lstFlags
            If counter = 1 Then
                tmpInt = tmpInt Or 1
                counter = 2
            Else
                tmpInt = tmpInt Or counter
                counter *= 2
            End If
        Next

        Return tmpInt
    End Function

    ''' <summary>
    ''' Convierte un entero en una lista de booleanos
    ''' </summary>
    ''' <param name="flags"></param>
    ''' <returns></returns>
    Public Function IntToFlags(ByVal flags As Integer) As List(Of Boolean)
        Dim counter As Integer = 1
        Dim tmpFlags As New List(Of Boolean)

        While counter <= 256
            If (flags And counter) = 0 Then
                tmpFlags.Add(False)
            Else
                tmpFlags.Add(True)
            End If
            counter *= 2
        End While

        Return tmpFlags
    End Function

    Public Function GetPrimerDiaMes(ByVal mes As Integer, ByVal anio As Integer) As Date
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

    Public Function GetUltimoDiaMes(ByVal mes As Integer, ByVal anio As Integer) As Date

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
        Dim tmpDate As Date
        Date.TryParse(tmpDiasMes.ToString & "/" & tmpMes & "/" & anio, tmpDate)
        Return tmpDate.Date
    End Function




    Public Function toSQLDateString(ByVal origDate As Date) As Integer

        Dim tmpDay As String = ""

        Dim sqlString As String = origDate.Year.ToString



        If origDate.Month < 10 Then
            tmpDay = "0" & origDate.Month.ToString
            sqlString &= tmpDay
        Else
            sqlString &= origDate.Month.ToString
        End If

        If origDate.Day < 10 Then
            tmpDay = "0" & origDate.Day.ToString
            sqlString &= tmpDay
        Else
            sqlString &= origDate.Day.ToString
        End If


        Return CInt(sqlString)
    End Function

    Public Function RedondeoPeso(ByVal valor As Decimal) As Decimal
        Dim redondeo As Decimal = (valor - Math.Truncate(valor)) * 100
        If redondeo <= 50 Then
            Return Math.Truncate(valor)
        Else
            Return Math.Truncate(valor) + 1
        End If
    End Function


    Public Function LetraToInt(ByVal letra As String) As Integer
        Return CInt(Convert.ToByte(letra.ToUpper.First) - 65)
    End Function

    Public Function TraducirCondicionFiscal(Optional comboIndex As Integer = -1, Optional dbIndex As Integer = -1) As Integer
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

    Public Function GetRubro(ByVal dbIndex As Integer) As String
        Select Case dbIndex
            Case 1
                Return "Comercios"
            Case 2
                Return "Servicios"
            Case 3
                Return "Industrias"
        End Select
    End Function


    Public Function ToISO8601(ByVal fecha As Date) As Integer
        Dim tmpMes As String = ""
        Dim tmpDia As String = ""

        If fecha.Month < 10 Then
            tmpMes = "0" & fecha.Month.ToString
        Else
            tmpMes = fecha.Month.ToString
        End If

        If fecha.Day < 10 Then
            tmpDia = "0" & fecha.Day.ToString
        Else
            tmpDia = fecha.Day.ToString
        End If

        Return CInt(fecha.Year.ToString & tmpMes & tmpDia)
    End Function

    Public Function fromISO8601(ByVal fecha As Integer) As String
        If fecha > 19000000 Then
            Dim tmpStr As String
            Dim strFecha As String = fecha.ToString
            tmpStr = $"{strFecha.Substring(6, 2)}/{strFecha.Substring(4, 2)}/{strFecha.Substring(0, 4)}"
            Return tmpStr
        Else
            Return "01/01/2000"
        End If
    End Function
    Public Function fromISO8601Date(ByVal fecha As Integer) As Date
        Dim datStr As String = fromISO8601(fecha)
        Dim isoDate As DateTime = DateTime.ParseExact(datStr, "dd/MM/yyyy", CultureInfo.InvariantCulture)
        Return isoDate
    End Function

    Public Function DescargarConstanciaInscripcion(ByVal cuit As String) As Boolean
        If Utils.ValidarCUIT(cuit) Then
            Dim dir As New Uri($"https://soa.afip.gob.ar/sr-padron/v1/constancia/{cuit}")

            ' FIX: Los genios de AFIP no renovaron el certificado y tira error al descargar las constancias
            System.Net.ServicePointManager.ServerCertificateValidationCallback =
              Function(se As Object,
              cert As System.Security.Cryptography.X509Certificates.X509Certificate,
              chain As System.Security.Cryptography.X509Certificates.X509Chain,
              sslerror As System.Net.Security.SslPolicyErrors) True
            Try
                My.Computer.Network.DownloadFile(dir, $"{My.Computer.FileSystem.SpecialDirectories.MyDocuments}\CONSTANCIA_{cuit}.pdf", Nothing, True, 500, True)
                System.Diagnostics.Process.Start($"{My.Computer.FileSystem.SpecialDirectories.MyDocuments}\CONSTANCIA_{cuit}.pdf")
                Return True
            Catch ex As Exception
            End Try

            Try
                My.Computer.Network.DownloadFile(dir, $"{My.Computer.FileSystem.SpecialDirectories.MyDocuments}\CONSTANCIA_{cuit}.pdf", Nothing, True, 500, True)
                System.Diagnostics.Process.Start($"{My.Computer.FileSystem.SpecialDirectories.MyDocuments}\CONSTANCIA_{cuit}.pdf")
                Return True
            Catch ex As Exception
                MsgBox("No se pudo finalizar la operación", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Descargar constancia de inscripción")
            End Try

            Return False

        Else
            Return False
        End If
    End Function

    Public Function ValidarCUIT(ByVal cuit As String) As Boolean
        If (cuit.Trim.Length <> 11) Or Not IsNumeric(cuit) Then Return False
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
    End Function

    Public Function GetCondicionFiscalString(ByVal condicionFiscal As Integer) As String
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

    Public Function CalcularSubtotal(ByVal precioUnitario As Decimal, ByVal porcentajeBonficacion As Decimal) As Decimal
        Dim total As Decimal = precioUnitario
        Return (total - CalcularPorcentaje(total, porcentajeBonficacion))
    End Function

    Public Function CalcularPorcentaje(ByVal total As Decimal, ByVal porcentaje As Decimal) As Decimal
        Dim resultado As Decimal = (porcentaje * total) / 100
        Return (resultado)
    End Function


    Public Function StringToFile(ByVal cadena As String, ByVal filePath As String) As Boolean
        Try
            Dim strFile As String = filePath
            Dim fileExists As Boolean = File.Exists(strFile)
            Using swC As New StreamWriter(File.Open(strFile, FileMode.OpenOrCreate))
                swC.Write(cadena)
            End Using

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function


    Public Function InjectSqle(ByVal origenSqle As SQLEngine, ByRef destinoSqle As SQLEngine)
        destinoSqle.RequireCredentials = origenSqle.RequireCredentials
        destinoSqle.Username = origenSqle.Username
        destinoSqle.Password = origenSqle.Password
        destinoSqle.dbType = origenSqle.dbType
        destinoSqle.Path = origenSqle.Path
        destinoSqle.DatabaseName = origenSqle.DatabaseName
        If origenSqle.IsStarted Then
            Return destinoSqle.ColdBoot()
        Else
            Return destinoSqle.Start()
        End If
    End Function
End Module
