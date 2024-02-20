Imports helix
Public Class ContratoCofre
    Public Property Sqle As New SQLEngine
    Public Property SearchResult As New List(Of ContratoCofre)

    Public Property Id As Long = 0
    Public Property Tipo As Byte = 0
    Public Property Numero As Integer = 0
    Public Property EsSocioId As Long = 0
    Public Property CofreLetra As String = ""
    Public Property CofreNumero As Integer = 0
    Public Property CofreTipo As Byte = 0
    Public Property Nombre As String = ""
    Public Property Modalidad As Byte = 0
    Public Property Conjunta1 As Boolean = False
    Public Property Conjunta2 As Boolean = False
    Public Property Conjunta3 As Boolean = False
    Public Property FechaContratacion As Date = #1/1/1970#
    Public Property FechaVencimiento As Date = #1/1/1970#
    Public Property Estado As Byte = 0
    Public Property RecibirInfo As Boolean = False
    Public Property ContactoCalle As String = ""
    Public Property ContactoCalleNum As String = ""
    Public Property ContactoCallePiso As String = ""
    Public Property ContactoCalleDepto As String = ""
    Public Property ContactoCP As String = ""
    Public Property ContactoCiudad As String = ""
    Public Property ContactoProvincia As Byte = 0
    Public Property ContactoTel As String = ""
    Public Property ContactoCel As String = ""
    Public Property ContactoMail As String = ""
    Public Property Deleted As Boolean = False
    Public Property Modificado As Date = Now

    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum


    Public Structure TABLA
        Const TABLA_NOMBRE As String = "contratoCofres"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const TIPO As String = TABLA_NOMBRE & "_tipo"
        Const NUMERO As String = TABLA_NOMBRE & "_numero"
        Const ES_SOCIO_ID As String = TABLA_NOMBRE & "_esSocioId"
        Const COFRE_LETRA As String = TABLA_NOMBRE & "_cofreLetra"
        Const COFRE_NUMERO As String = TABLA_NOMBRE & "_cofreNumero"
        Const COFRE_TIPO As String = TABLA_NOMBRE & "_cofreTipo"
        Const NOMBRE As String = TABLA_NOMBRE & "_nombre"
        Const MODALIDAD As String = TABLA_NOMBRE & "_modalidad"
        Const CONJUNTA_1 As String = TABLA_NOMBRE & "_conjunta1"
        Const CONJUNTA_2 As String = TABLA_NOMBRE & "_conjunta2"
        Const CONJUNTA_3 As String = TABLA_NOMBRE & "_conjunta3"
        Const FECHA_CONTRATACION As String = TABLA_NOMBRE & "_fechaContratacion"
        Const FECHA_VENCIMIENTO As String = TABLA_NOMBRE & "_fechaVencimiento"
        Const ESTADO As String = TABLA_NOMBRE & "_estado"
        Const RECIBIR_INFO As String = TABLA_NOMBRE & "_recibirInfo"
        Const CONTACTO_CALLE As String = TABLA_NOMBRE & "_contactoCalle"
        Const CONTACTO_CALLE_NUM As String = TABLA_NOMBRE & "_contactoCalleNum"
        Const CONTACTO_CALLE_PISO As String = TABLA_NOMBRE & "_contactoCallePiso"
        Const CONTACTO_CALLE_DEPTO As String = TABLA_NOMBRE & "_contactoCalleDepto"
        Const CONTACTO_C_P As String = TABLA_NOMBRE & "_contactoCP"
        Const CONTACTO_CIUDAD As String = TABLA_NOMBRE & "_contactoCiudad"
        Const CONTACTO_PROVINCIA As String = TABLA_NOMBRE & "_contactoProvincia"
        Const CONTACTO_TEL As String = TABLA_NOMBRE & "_contactoTel"
        Const CONTACTO_CEL As String = TABLA_NOMBRE & "_contactoCel"
        Const CONTACTO_MAIL As String = TABLA_NOMBRE & "_contactoMail"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & TIPO & ", " & NUMERO & ", " & ES_SOCIO_ID & ", " & COFRE_LETRA & ", " & COFRE_NUMERO & ", " & COFRE_TIPO & ", " & NOMBRE & ", " & MODALIDAD & ", " & CONJUNTA_1 & ", " & CONJUNTA_2 & ", " & CONJUNTA_3 & ", " & FECHA_CONTRATACION & ", " & FECHA_VENCIMIENTO & ", " & ESTADO & ", " & RECIBIR_INFO & ", " & CONTACTO_CALLE & ", " & CONTACTO_CALLE_NUM & ", " & CONTACTO_CALLE_PISO & ", " & CONTACTO_CALLE_DEPTO & ", " & CONTACTO_C_P & ", " & CONTACTO_CIUDAD & ", " & CONTACTO_PROVINCIA & ", " & CONTACTO_TEL & ", " & CONTACTO_CEL & ", " & CONTACTO_MAIL & ", " & DELETED & ", " & MODIFICADO
    End Structure



    Public Sub New()
    End Sub


    Public Sub New(ByVal sqle As SQLEngine)
        Me.Sqle.RequireCredentials = sqle.RequireCredentials
        Me.Sqle.Username = sqle.Username
        Me.Sqle.Password = sqle.Password
        Me.Sqle.dbType = sqle.dbType
        Me.Sqle.Path = sqle.Path
        Me.Sqle.DatabaseName = sqle.DatabaseName
        If sqle.IsStarted Then
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
            .SimpleSearch(TABLA.ID, SQLEngineQuery.OperatorCriteria.Igual, myID)
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    Tipo = CByte(.GetQueryData(TABLA.TIPO))
                    Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    EsSocioId = CLng(.GetQueryData(TABLA.ES_SOCIO_ID))
                    CofreLetra = CStr(.GetQueryData(TABLA.COFRE_LETRA))
                    CofreNumero = CInt(.GetQueryData(TABLA.COFRE_NUMERO))
                    CofreTipo = CByte(.GetQueryData(TABLA.COFRE_TIPO))
                    Nombre = CStr(.GetQueryData(TABLA.NOMBRE))
                    Modalidad = CByte(.GetQueryData(TABLA.MODALIDAD))
                    Conjunta1 = CBool(.GetQueryData(TABLA.CONJUNTA_1))
                    Conjunta2 = CBool(.GetQueryData(TABLA.CONJUNTA_2))
                    Conjunta3 = CBool(.GetQueryData(TABLA.CONJUNTA_3))
                    FechaContratacion = CDate(.GetQueryData(TABLA.FECHA_CONTRATACION))
                    FechaVencimiento = CDate(.GetQueryData(TABLA.FECHA_VENCIMIENTO))
                    Estado = CByte(.GetQueryData(TABLA.ESTADO))
                    RecibirInfo = CBool(.GetQueryData(TABLA.RECIBIR_INFO))
                    ContactoCalle = CStr(.GetQueryData(TABLA.CONTACTO_CALLE))
                    ContactoCalleNum = CStr(.GetQueryData(TABLA.CONTACTO_CALLE_NUM))
                    ContactoCallePiso = CStr(.GetQueryData(TABLA.CONTACTO_CALLE_PISO))
                    ContactoCalleDepto = CStr(.GetQueryData(TABLA.CONTACTO_CALLE_DEPTO))
                    ContactoCP = CStr(.GetQueryData(TABLA.CONTACTO_C_P))
                    ContactoCiudad = CStr(.GetQueryData(TABLA.CONTACTO_CIUDAD))
                    ContactoProvincia = CByte(.GetQueryData(TABLA.CONTACTO_PROVINCIA))
                    ContactoTel = CStr(.GetQueryData(TABLA.CONTACTO_TEL))
                    ContactoCel = CStr(.GetQueryData(TABLA.CONTACTO_CEL))
                    ContactoMail = CStr(.GetQueryData(TABLA.CONTACTO_MAIL))
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
                    .AddColumnValue(TABLA.TIPO, Tipo)
                    .AddColumnValue(TABLA.NUMERO, Numero)
                    .AddColumnValue(TABLA.ES_SOCIO_ID, EsSocioId)
                    .AddColumnValue(TABLA.COFRE_LETRA, CofreLetra)
                    .AddColumnValue(TABLA.COFRE_NUMERO, CofreNumero)
                    .AddColumnValue(TABLA.COFRE_TIPO, CofreTipo)
                    .AddColumnValue(TABLA.NOMBRE, Nombre)
                    .AddColumnValue(TABLA.MODALIDAD, Modalidad)
                    .AddColumnValue(TABLA.CONJUNTA_1, Conjunta1)
                    .AddColumnValue(TABLA.CONJUNTA_2, Conjunta2)
                    .AddColumnValue(TABLA.CONJUNTA_3, Conjunta3)
                    .AddColumnValue(TABLA.FECHA_CONTRATACION, FechaContratacion)
                    .AddColumnValue(TABLA.FECHA_VENCIMIENTO, FechaVencimiento)
                    .AddColumnValue(TABLA.ESTADO, Estado)
                    .AddColumnValue(TABLA.RECIBIR_INFO, RecibirInfo)
                    .AddColumnValue(TABLA.CONTACTO_CALLE, ContactoCalle)
                    .AddColumnValue(TABLA.CONTACTO_CALLE_NUM, ContactoCalleNum)
                    .AddColumnValue(TABLA.CONTACTO_CALLE_PISO, ContactoCallePiso)
                    .AddColumnValue(TABLA.CONTACTO_CALLE_DEPTO, ContactoCalleDepto)
                    .AddColumnValue(TABLA.CONTACTO_C_P, ContactoCP)
                    .AddColumnValue(TABLA.CONTACTO_CIUDAD, ContactoCiudad)
                    .AddColumnValue(TABLA.CONTACTO_PROVINCIA, ContactoProvincia)
                    .AddColumnValue(TABLA.CONTACTO_TEL, ContactoTel)
                    .AddColumnValue(TABLA.CONTACTO_CEL, ContactoCel)
                    .AddColumnValue(TABLA.CONTACTO_MAIL, ContactoMail)
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
                With Sqle.Update
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.TIPO, Tipo)
                    .AddColumnValue(TABLA.NUMERO, Numero)
                    .AddColumnValue(TABLA.ES_SOCIO_ID, EsSocioId)
                    .AddColumnValue(TABLA.COFRE_LETRA, CofreLetra)
                    .AddColumnValue(TABLA.COFRE_NUMERO, CofreNumero)
                    .AddColumnValue(TABLA.COFRE_TIPO, CofreTipo)
                    .AddColumnValue(TABLA.NOMBRE, Nombre)
                    .AddColumnValue(TABLA.MODALIDAD, Modalidad)
                    .AddColumnValue(TABLA.CONJUNTA_1, Conjunta1)
                    .AddColumnValue(TABLA.CONJUNTA_2, Conjunta2)
                    .AddColumnValue(TABLA.CONJUNTA_3, Conjunta3)
                    .AddColumnValue(TABLA.FECHA_CONTRATACION, FechaContratacion)
                    .AddColumnValue(TABLA.FECHA_VENCIMIENTO, FechaVencimiento)
                    .AddColumnValue(TABLA.ESTADO, Estado)
                    .AddColumnValue(TABLA.RECIBIR_INFO, RecibirInfo)
                    .AddColumnValue(TABLA.CONTACTO_CALLE, ContactoCalle)
                    .AddColumnValue(TABLA.CONTACTO_CALLE_NUM, ContactoCalleNum)
                    .AddColumnValue(TABLA.CONTACTO_CALLE_PISO, ContactoCallePiso)
                    .AddColumnValue(TABLA.CONTACTO_CALLE_DEPTO, ContactoCalleDepto)
                    .AddColumnValue(TABLA.CONTACTO_C_P, ContactoCP)
                    .AddColumnValue(TABLA.CONTACTO_CIUDAD, ContactoCiudad)
                    .AddColumnValue(TABLA.CONTACTO_PROVINCIA, ContactoProvincia)
                    .AddColumnValue(TABLA.CONTACTO_TEL, ContactoTel)
                    .AddColumnValue(TABLA.CONTACTO_CEL, ContactoCel)
                    .AddColumnValue(TABLA.CONTACTO_MAIL, ContactoMail)
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
                    Dim tmp As New ContratoCofre
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.Tipo = CByte(.GetQueryData(TABLA.TIPO))
                    tmp.Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    tmp.EsSocioId = CLng(.GetQueryData(TABLA.ES_SOCIO_ID))
                    tmp.CofreLetra = CStr(.GetQueryData(TABLA.COFRE_LETRA))
                    tmp.CofreNumero = CInt(.GetQueryData(TABLA.COFRE_NUMERO))
                    tmp.CofreTipo = CByte(.GetQueryData(TABLA.COFRE_TIPO))
                    tmp.Nombre = CStr(.GetQueryData(TABLA.NOMBRE))
                    tmp.Modalidad = CByte(.GetQueryData(TABLA.MODALIDAD))
                    tmp.Conjunta1 = CBool(.GetQueryData(TABLA.CONJUNTA_1))
                    tmp.Conjunta2 = CBool(.GetQueryData(TABLA.CONJUNTA_2))
                    tmp.Conjunta3 = CBool(.GetQueryData(TABLA.CONJUNTA_3))
                    tmp.FechaContratacion = CDate(.GetQueryData(TABLA.FECHA_CONTRATACION))
                    tmp.FechaVencimiento = CDate(.GetQueryData(TABLA.FECHA_VENCIMIENTO))
                    tmp.Estado = CByte(.GetQueryData(TABLA.ESTADO))
                    tmp.RecibirInfo = CBool(.GetQueryData(TABLA.RECIBIR_INFO))
                    tmp.ContactoCalle = CStr(.GetQueryData(TABLA.CONTACTO_CALLE))
                    tmp.ContactoCalleNum = CStr(.GetQueryData(TABLA.CONTACTO_CALLE_NUM))
                    tmp.ContactoCallePiso = CStr(.GetQueryData(TABLA.CONTACTO_CALLE_PISO))
                    tmp.ContactoCalleDepto = CStr(.GetQueryData(TABLA.CONTACTO_CALLE_DEPTO))
                    tmp.ContactoCP = CStr(.GetQueryData(TABLA.CONTACTO_C_P))
                    tmp.ContactoCiudad = CStr(.GetQueryData(TABLA.CONTACTO_CIUDAD))
                    'tmp.ContactoProvincia = CByte(.GetQueryData(TABLA.CONTACTO_PROVINCIA))
                    tmp.ContactoTel = CStr(.GetQueryData(TABLA.CONTACTO_TEL))
                    tmp.ContactoCel = CStr(.GetQueryData(TABLA.CONTACTO_CEL))
                    tmp.ContactoMail = CStr(.GetQueryData(TABLA.CONTACTO_MAIL))
                    tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))
                    SearchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function
End Class


