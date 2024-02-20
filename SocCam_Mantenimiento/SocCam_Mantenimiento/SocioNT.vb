Imports helix

Public Class SocioNT
    Public Property Sqle As New SQLEngine
    Public Property SearchResult As New List(Of SocioNT)

    Public Property Id As Long = 0
    Public Property Nombre As String = ""
    Public Property Apellido As String = ""
    Public Property Nacionalidad As String = ""
    Public Property Dni As String = ""
    Public Property FechaNacimiento As Date = #1/1/1970#
    Public Property Cuit As String = ""
    Public Property Mail As String = ""
    Public Property Firma As String = ""
    Public Property TipoEmpresa As String = ""
    Public Property Domicilio As String = ""
    Public Property Telefono As String = ""
    Public Property Celular As String = ""
    Public Property OtroTelefono As String = ""
    Public Property TipoSocio As Long = 0
    Public Property Numero As Integer = 0
    Public Property FechaAprobacion As Date = #1/1/1970#
    Public Property Acta As Integer = 0
    Public Property Padrino1 As Long = 0
    Public Property Padrino2 As Long = 0
    Public Property Sector As Integer = 0
    Public Property Deleted As Boolean = False
    Public Property TieneCajaSeguridad As Boolean = False
    Public Property Gestion As Long = 0
    Public Property Segmento As Long = 0
    Public Property Rubro As Long = 0
    Public Property Localidad As Long = 0
    Public Property Habilitacion As Integer = 0
    Public Property CondicionFiscal As Integer = 0
    Public Property TarjetaEntregada As Boolean = False
    Public Property TarjetaFechaEntrega As Integer = 0
    Public Property Campania As Long = 0
    Public Property EnviarMail As Boolean = False
    Public Property Estado As Integer = 0
    Public Property Modificado As Date = Now

    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum

    Public Structure TABLA
        Const TABLA_NOMBRE As String = "socio"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const NOMBRE As String = TABLA_NOMBRE & "_nombre"
        Const APELLIDO As String = TABLA_NOMBRE & "_apellido"
        Const NACIONALIDAD As String = TABLA_NOMBRE & "_nacionalidad"
        Const DNI As String = TABLA_NOMBRE & "_dni"
        Const FECHA_NACIMIENTO As String = TABLA_NOMBRE & "_fechaNacimiento"
        Const CUIT As String = TABLA_NOMBRE & "_cuit"
        Const MAIL As String = TABLA_NOMBRE & "_mail"
        Const FIRMA As String = TABLA_NOMBRE & "_firma"
        Const TIPO_EMPRESA As String = TABLA_NOMBRE & "_tipoEmpresa"
        Const DOMICILIO As String = TABLA_NOMBRE & "_domicilio"
        Const TELEFONO As String = TABLA_NOMBRE & "_telefono"
        Const CELULAR As String = TABLA_NOMBRE & "_celular"
        Const OTRO_TELEFONO As String = TABLA_NOMBRE & "_otroTelefono"
        Const TIPO_SOCIO As String = TABLA_NOMBRE & "_tipoSocio"
        Const NUMERO As String = TABLA_NOMBRE & "_numero"
        Const FECHA_APROBACION As String = TABLA_NOMBRE & "_fechaAprobacion"
        Const ACTA As String = TABLA_NOMBRE & "_acta"
        Const PADRINO_1 As String = TABLA_NOMBRE & "_padrino1"
        Const PADRINO_2 As String = TABLA_NOMBRE & "_padrino2"
        Const SECTOR As String = TABLA_NOMBRE & "_sector"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const TIENE_CAJA_SEGURIDAD As String = TABLA_NOMBRE & "_tieneCajaSeguridad"
        Const GESTION As String = TABLA_NOMBRE & "_gestion"
        Const SEGMENTO As String = TABLA_NOMBRE & "_segmento"
        Const RUBRO As String = TABLA_NOMBRE & "_rubro"
        Const LOCALIDAD As String = TABLA_NOMBRE & "_localidad"
        Const HABILITACION As String = TABLA_NOMBRE & "_habilitacion"
        Const CONDICION_FISCAL As String = TABLA_NOMBRE & "_condicionFiscal"
        Const _TARJETA_ENTREGADA As String = TABLA_NOMBRE & "_TarjetaEntregada"
        Const _TARJETA_FECHA_ENTREGA As String = TABLA_NOMBRE & "_TarjetaFechaEntrega"
        Const _CAMPANIA As String = TABLA_NOMBRE & "_Campanias"
        Const _ENVIAR_MAIL As String = TABLA_NOMBRE & "_EnviarMail"
        Const ESTADO As String = TABLA_NOMBRE & "_estado"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & NOMBRE & ", " & APELLIDO & ", " & NACIONALIDAD & ", " & DNI & ", " & FECHA_NACIMIENTO & ", " & CUIT & ", " & MAIL & ", " & FIRMA & ", " & TIPO_EMPRESA & ", " & DOMICILIO & ", " & TELEFONO & ", " & CELULAR & ", " & OTRO_TELEFONO & ", " & TIPO_SOCIO & ", " & NUMERO & ", " & FECHA_APROBACION & ", " & ACTA & ", " & PADRINO_1 & ", " & PADRINO_2 & ", " & SECTOR & ", " & DELETED & ", " & TIENE_CAJA_SEGURIDAD & ", " & GESTION & ", " & SEGMENTO & ", " & RUBRO & ", " & HABILITACION & ", " & CONDICION_FISCAL & ", " & _TARJETA_ENTREGADA & ", " & _TARJETA_FECHA_ENTREGA & ", " & _CAMPANIA & ", " & _ENVIAR_MAIL & ", " & ESTADO & ", " & MODIFICADO & ", " & LOCALIDAD
    End Structure

    Public Sub New()
    End Sub

    Public Sub New(ByVal iSqle As SQLEngine)
        Me.Sqle.RequireCredentials = iSqle.RequireCredentials
        Me.Sqle.Username = iSqle.Username
        Me.Sqle.Password = iSqle.Password
        Me.Sqle.dbType = iSqle.dbType
        Me.Sqle.Path = iSqle.Path
        Me.Sqle.DatabaseName = iSqle.DatabaseName
        If iSqle.IsStarted Then
            Me.Sqle.ColdBoot()
        Else
            Me.Sqle.Start()
        End If
    End Sub

    Public Function LoadMe(ByVal myID As Integer) As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)

            .WHEREstring = $"{TABLA.ID} = { .p(myID)}"

            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    Nombre = CStr(.GetQueryData(TABLA.NOMBRE))
                    Apellido = CStr(.GetQueryData(TABLA.APELLIDO))
                    Nacionalidad = CStr(.GetQueryData(TABLA.NACIONALIDAD))
                    Dni = CStr(.GetQueryData(TABLA.DNI))
                    FechaNacimiento = CDate(.GetQueryData(TABLA.FECHA_NACIMIENTO))
                    Cuit = CStr(.GetQueryData(TABLA.CUIT))
                    Mail = CStr(.GetQueryData(TABLA.MAIL))
                    Firma = CStr(.GetQueryData(TABLA.FIRMA))
                    TipoEmpresa = CStr(.GetQueryData(TABLA.TIPO_EMPRESA))
                    Domicilio = CStr(.GetQueryData(TABLA.DOMICILIO))
                    Telefono = CStr(.GetQueryData(TABLA.TELEFONO))
                    Celular = CStr(.GetQueryData(TABLA.CELULAR))
                    OtroTelefono = CStr(.GetQueryData(TABLA.OTRO_TELEFONO))
                    TipoSocio = CLng(.GetQueryData(TABLA.TIPO_SOCIO))
                    Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    FechaAprobacion = CDate(.GetQueryData(TABLA.FECHA_APROBACION))
                    If .GetQueryData(TABLA.ACTA).ToString.Length <> 0 Then
                        Acta = CInt(.GetQueryData(TABLA.ACTA))
                    Else
                        Acta = 0
                    End If
                    Padrino1 = CLng(.GetQueryData(TABLA.PADRINO_1))
                    Padrino2 = CLng(.GetQueryData(TABLA.PADRINO_2))
                    Sector = CInt(.GetQueryData(TABLA.SECTOR))
                    Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    TieneCajaSeguridad = CBool(.GetQueryData(TABLA.TIENE_CAJA_SEGURIDAD))
                    Gestion = CLng(.GetQueryData(TABLA.GESTION))
                    Segmento = CLng(.GetQueryData(TABLA.SEGMENTO))
                    Rubro = CLng(.GetQueryData(TABLA.RUBRO))
                    Habilitacion = CInt(.GetQueryData(TABLA.HABILITACION))
                    CondicionFiscal = CInt(.GetQueryData(TABLA.CONDICION_FISCAL))
                    If CStr(.GetQueryData(TABLA._TARJETA_ENTREGADA)).Length > 0 Then
                        TarjetaEntregada = CBool(.GetQueryData(TABLA._TARJETA_ENTREGADA))
                        TarjetaFechaEntrega = CInt(.GetQueryData(TABLA._TARJETA_FECHA_ENTREGA))
                    Else
                        TarjetaEntregada = False
                        TarjetaFechaEntrega = Nothing
                    End If

                    If CStr(.GetQueryData(TABLA._CAMPANIA)).Length > 0 Then
                        Campania = CLng(.GetQueryData(TABLA._CAMPANIA))
                    Else
                        Campania = 0
                    End If

                    If CStr(.GetQueryData(TABLA.LOCALIDAD)).Length > 0 Then
                        Localidad = CLng(.GetQueryData(TABLA.LOCALIDAD))
                    Else
                        Localidad = 0
                    End If



                    EnviarMail = CBool(.GetQueryData(TABLA._ENVIAR_MAIL))
                    Estado = CInt(.GetQueryData(TABLA.ESTADO))
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

    Public Function LoadAll() As Integer
        Me.QuickSearch(TABLA.DELETED, SQLEngineQuery.OperatorCriteria.Igual, False)
        Return SearchResult.Count
    End Function

    Public Function LoadAll(ByRef dt As DataTable) As Boolean
        With Sqle.Query
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
                With Sqle.Insert
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.NOMBRE, Nombre)
                    .AddColumnValue(TABLA.APELLIDO, Apellido)
                    .AddColumnValue(TABLA.NACIONALIDAD, Nacionalidad)
                    .AddColumnValue(TABLA.DNI, Dni)
                    .AddColumnValue(TABLA.FECHA_NACIMIENTO, FechaNacimiento)
                    .AddColumnValue(TABLA.CUIT, Cuit)
                    .AddColumnValue(TABLA.MAIL, Mail)
                    .AddColumnValue(TABLA.FIRMA, Firma)
                    .AddColumnValue(TABLA.TIPO_EMPRESA, TipoEmpresa)
                    .AddColumnValue(TABLA.DOMICILIO, Domicilio)
                    .AddColumnValue(TABLA.TELEFONO, Telefono)
                    .AddColumnValue(TABLA.CELULAR, Celular)
                    .AddColumnValue(TABLA.OTRO_TELEFONO, OtroTelefono)
                    .AddColumnValue(TABLA.TIPO_SOCIO, TipoSocio)
                    .AddColumnValue(TABLA.NUMERO, Numero)
                    .AddColumnValue(TABLA.FECHA_APROBACION, FechaAprobacion)
                    .AddColumnValue(TABLA.ACTA, Acta)
                    .AddColumnValue(TABLA.PADRINO_1, Padrino1)
                    .AddColumnValue(TABLA.PADRINO_2, Padrino2)
                    .AddColumnValue(TABLA.SECTOR, Sector)
                    .AddColumnValue(TABLA.DELETED, Deleted)
                    .AddColumnValue(TABLA.TIENE_CAJA_SEGURIDAD, TieneCajaSeguridad)
                    .AddColumnValue(TABLA.GESTION, Gestion)
                    .AddColumnValue(TABLA.SEGMENTO, Segmento)
                    .AddColumnValue(TABLA.RUBRO, Rubro)
                    .AddColumnValue(TABLA.HABILITACION, Habilitacion)
                    .AddColumnValue(TABLA.CONDICION_FISCAL, CondicionFiscal)
                    .AddColumnValue(TABLA._TARJETA_ENTREGADA, TarjetaEntregada)
                    .AddColumnValue(TABLA._TARJETA_FECHA_ENTREGA, TarjetaFechaEntrega)
                    .AddColumnValue(TABLA._CAMPANIA, Campania)
                    .AddColumnValue(TABLA._ENVIAR_MAIL, EnviarMail)
                    .AddColumnValue(TABLA.ESTADO, Estado)
                    .AddColumnValue(TABLA.LOCALIDAD, Localidad)
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
                With Sqle.Update
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.NOMBRE, Nombre)
                    .AddColumnValue(TABLA.APELLIDO, Apellido)
                    .AddColumnValue(TABLA.NACIONALIDAD, Nacionalidad)
                    .AddColumnValue(TABLA.DNI, Dni)
                    .AddColumnValue(TABLA.FECHA_NACIMIENTO, FechaNacimiento)
                    .AddColumnValue(TABLA.CUIT, Cuit)
                    .AddColumnValue(TABLA.MAIL, Mail)
                    .AddColumnValue(TABLA.FIRMA, Firma)
                    .AddColumnValue(TABLA.TIPO_EMPRESA, TipoEmpresa)
                    .AddColumnValue(TABLA.DOMICILIO, Domicilio)
                    .AddColumnValue(TABLA.TELEFONO, Telefono)
                    .AddColumnValue(TABLA.CELULAR, Celular)
                    .AddColumnValue(TABLA.OTRO_TELEFONO, OtroTelefono)
                    .AddColumnValue(TABLA.TIPO_SOCIO, TipoSocio)
                    .AddColumnValue(TABLA.NUMERO, Numero)
                    .AddColumnValue(TABLA.FECHA_APROBACION, FechaAprobacion)
                    .AddColumnValue(TABLA.ACTA, Acta)
                    .AddColumnValue(TABLA.PADRINO_1, Padrino1)
                    .AddColumnValue(TABLA.PADRINO_2, Padrino2)
                    .AddColumnValue(TABLA.SECTOR, Sector)
                    .AddColumnValue(TABLA.DELETED, Deleted)
                    .AddColumnValue(TABLA.TIENE_CAJA_SEGURIDAD, TieneCajaSeguridad)
                    .AddColumnValue(TABLA.GESTION, Gestion)
                    .AddColumnValue(TABLA.SEGMENTO, Segmento)
                    .AddColumnValue(TABLA.RUBRO, Rubro)
                    .AddColumnValue(TABLA.HABILITACION, Habilitacion)
                    .AddColumnValue(TABLA.CONDICION_FISCAL, CondicionFiscal)
                    .AddColumnValue(TABLA._TARJETA_ENTREGADA, TarjetaEntregada)
                    .AddColumnValue(TABLA._TARJETA_FECHA_ENTREGA, TarjetaFechaEntrega)
                    .AddColumnValue(TABLA._CAMPANIA, Campania)
                    .AddColumnValue(TABLA._ENVIAR_MAIL, EnviarMail)
                    .AddColumnValue(TABLA.LOCALIDAD, Localidad)
                    .AddColumnValue(TABLA.ESTADO, Estado)
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
            With Sqle.Delete
                .Reset()
                .TableName = TABLA.TABLA_NOMBRE
                .SimpleSearch(TABLA.ID, SQLEngineUpdate.OperatorCriteria.Igual, Id)
                Return .Delete
            End With
        Else
            With Sqle.Update
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
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(columna.ToString, operador, value)
            If .Query() Then
                Me.SearchResult.Clear()
                While .QueryRead
                    Dim tmp As New SocioNT
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.Nombre = CStr(.GetQueryData(TABLA.NOMBRE))
                    tmp.Apellido = CStr(.GetQueryData(TABLA.APELLIDO))
                    tmp.Nacionalidad = CStr(.GetQueryData(TABLA.NACIONALIDAD))
                    tmp.Dni = CStr(.GetQueryData(TABLA.DNI))
                    tmp.FechaNacimiento = CDate(.GetQueryData(TABLA.FECHA_NACIMIENTO))
                    tmp.Cuit = CStr(.GetQueryData(TABLA.CUIT))
                    tmp.Mail = CStr(.GetQueryData(TABLA.MAIL))
                    tmp.Firma = CStr(.GetQueryData(TABLA.FIRMA))
                    tmp.TipoEmpresa = CStr(.GetQueryData(TABLA.TIPO_EMPRESA))
                    tmp.Domicilio = CStr(.GetQueryData(TABLA.DOMICILIO))
                    tmp.Telefono = CStr(.GetQueryData(TABLA.TELEFONO))
                    tmp.Celular = CStr(.GetQueryData(TABLA.CELULAR))
                    tmp.OtroTelefono = CStr(.GetQueryData(TABLA.OTRO_TELEFONO))
                    tmp.TipoSocio = CLng(.GetQueryData(TABLA.TIPO_SOCIO))
                    tmp.Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    tmp.FechaAprobacion = CDate(.GetQueryData(TABLA.FECHA_APROBACION))
                    tmp.Acta = CInt(.GetQueryData(TABLA.ACTA))
                    tmp.Padrino1 = CLng(.GetQueryData(TABLA.PADRINO_1))
                    tmp.Padrino2 = CLng(.GetQueryData(TABLA.PADRINO_2))
                    tmp.Sector = CInt(.GetQueryData(TABLA.SECTOR))
                    tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    tmp.TieneCajaSeguridad = CBool(.GetQueryData(TABLA.TIENE_CAJA_SEGURIDAD))
                    tmp.Gestion = CLng(.GetQueryData(TABLA.GESTION))
                    tmp.Segmento = CLng(.GetQueryData(TABLA.SEGMENTO))
                    tmp.Rubro = CLng(.GetQueryData(TABLA.RUBRO))
                    tmp.Localidad = CLng(.GetQueryData(TABLA.LOCALIDAD))
                    tmp.Habilitacion = CInt(.GetQueryData(TABLA.HABILITACION))
                    tmp.CondicionFiscal = CInt(.GetQueryData(TABLA.CONDICION_FISCAL))
                    If CStr(.GetQueryData(TABLA._TARJETA_ENTREGADA)).Length <> 0 Then
                        tmp.TarjetaEntregada = CBool(.GetQueryData(TABLA._TARJETA_ENTREGADA))
                    End If
                    If CStr(.GetQueryData(TABLA._TARJETA_FECHA_ENTREGA)).Length <> 0 Then
                        tmp.TarjetaFechaEntrega = CInt(.GetQueryData(TABLA._TARJETA_FECHA_ENTREGA))
                    End If

                    If CStr(.GetQueryData(TABLA._CAMPANIA)).Length <> 0 Then
                        tmp.Campania = CLng(.GetQueryData(TABLA._CAMPANIA))
                    End If

                    If CStr(.GetQueryData(TABLA._ENVIAR_MAIL)).Length <> 0 Then
                        tmp.EnviarMail = CBool(.GetQueryData(TABLA._ENVIAR_MAIL))
                    End If

                    tmp.Estado = CInt(.GetQueryData(TABLA.ESTADO))
                    tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))
                    SearchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function



End Class