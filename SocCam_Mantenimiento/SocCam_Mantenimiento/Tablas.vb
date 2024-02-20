''' <summary>
''' Estructuras de todas las tablas
''' </summary>
''' <remarks></remarks>
Public Module Tablas
    Public Structure TABLA_SOCIO
        Const TABLA_NOMBRE As String = "socio"
        Const ID As String = "socio_id"
        Const NOMBRE As String = "socio_nombre"
        Const APELLIDO As String = "socio_apellido"
        Const NACIONALIDAD As String = "socio_nacionalidad"
        Const DNI As String = "socio_dni"
        Const FECHA_NACIMIENTO As String = "socio_fechaNacimiento"
        Const CUIT As String = "socio_cuit"
        Const MAIL As String = "socio_mail"
        Const FIRMA As String = "socio_firma"
        Const TIPO_EMPRESA As String = "socio_tipoEmpresa"
        Const DOMICILIO As String = "socio_domicilio"
        Const LOCALIDAD As String = "socio_localidad"
        Const TELEFONO As String = "socio_telefono"
        Const CELULAR As String = "socio_celular"
        Const OTRO_TELEFONO As String = "socio_otroTelefono"
        Const TIPO_SOCIO As String = "socio_tipoSocio"
        Const NUMERO As String = "socio_numero"
        Const FECHA_APROBACION As String = "socio_fechaAprobacion"
        Const ACTA As String = "socio_acta"
        Const SOCIO_PADRINO1 As String = "socio_padrino1"
        Const SOCIO_PADRINO2 As String = "socio_padrino2"
        Const SECTOR As String = "socio_sector"
        Const CAJA_SEGURIDAD As String = "socio_tieneCajaSeguridad"
        Const MOTIVO_BAJA As String = "socio_motivoBaja"
        Const GESTION As String = "socio_gestion"
        Const SEGMENTO As String = "socio_segmento"
        Const RUBRO As String = "socio_rubro"
        Const HABILITACION As String = "socio_habilitacion"
        Const CONDICION_FISCAL As String = "socio_condicionFiscal"
        Const ESTADO As String = "socio_estado"
        Const TARJETA_ENTREGADA As String = "socio_tarjetaEntregada"
        Const TARJETA_FECHA_ENTREGA As String = "socio_tarjetaFechaEntrega"
        Const CAMPANIA As String = "socio_Campanias"
        Const DELETED As String = "socio_deleted"
        Const MODIFICADO As String = "socio_modificado"
        Const TIPO As String = "tipo"
        Const ENVIAR_MAIL As String = "socio_enviarMail"

        Const ALL As String = ID & "," & NOMBRE & "," & APELLIDO & "," & NACIONALIDAD & "," & DNI & "," & FECHA_NACIMIENTO & "," & FECHA_NACIMIENTO & "," &
                              CUIT & "," & MAIL & "," & FIRMA & "," & TIPO_EMPRESA & "," & DOMICILIO & "," & TELEFONO & "," & CELULAR & "," & OTRO_TELEFONO & "," & TIPO_SOCIO & "," &
                              NUMERO & "," & FECHA_APROBACION & "," & ACTA & "," & SOCIO_PADRINO1 & "," & SOCIO_PADRINO2 & "," & SECTOR & "," & CAJA_SEGURIDAD & "," &
                              ESTADO & "," & MOTIVO_BAJA & "," & GESTION & "," & SEGMENTO & "," & RUBRO & "," & HABILITACION & "," & CONDICION_FISCAL & "," & LOCALIDAD & "," & DELETED & "," & MODIFICADO & "," &
                              TARJETA_ENTREGADA & "," & TARJETA_FECHA_ENTREGA & "," & CAMPANIA & "," & ENVIAR_MAIL

    End Structure


    Public Structure TABLA_TIPO_SOCIO
        Const TABLA_NOMBRE As String = "tipoSocio"
        Const ID As String = "tipoSocio_id"
        Const TIPO As String = "tipoSocio_tipo"
        Const NOMBRE As String = "tipoSocio_nombre"
        Const IMPORTE As String = "tipoSocio_importe"
        Const PERIODICIDAD As String = "tipoSocio_periodicidad"
        Const DELETED As String = "tipoSocio_deleted"
        Const MODIFICADO As String = "tipoSocio_modificado"
        Const CATEGORIA As String = "tipoSocio_categoria"
        Const ALL As String = ID & ", " & TIPO & ", " & NOMBRE & ", " & IMPORTE & ", " & PERIODICIDAD & ", " & DELETED & ", " & MODIFICADO & ", " & CATEGORIA
    End Structure

    ''' <summary>
    ''' Tabla de pago de socios
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure TABLA_PAGO_SOCIOS
        Const TABLA_NOMBRE As String = "pagosSocios"
        Const ID As String = "pagosSocios_id"
        Const SOCIO As String = "pagosSocios_socio"
        Const PLAN As String = "pagosSocios_plan"
        Const ANIO As String = "pagosSocios_anio"
        Const PERIODO As String = "pagosSocios_periodo"
        Const PERIODICIDAD As String = "pagosSocios_periodicidad"
        Const FECHA_VENCIMIENTO As String = "pagosSocios_fechaVencimiento"
        Const FECHA_PAGO As String = "pagosSocios_fechaPago"
        Const OBSERVACIONES As String = "pagosSocios_observaciones"
        Const ESTADO As String = "pagosSocios_estado"
        Const COBRADOR As String = "pagosSocios_cobrador"
        Const MONTO As String = "pagosSocios_monto"

        Const DELETED As String = "pagosSocios_deleted"
        Const MODIFICADO As String = "pagosSocios_modificado"

        Const OPERACION As String = "pagosSocios_operacion"

        Const BUSQUEDA_VENCIMIENTO As String = "pagosSocios_crx_fechaVencimiento"
        Const BUSQUEDA_PAGO As String = "pagosSocios_crx_fechaPago"
        Const RECIBO_ID As String = "recibo_id"

        Const MOVIMIENTO_CC As String = "pagosSocios_movimiento_cc"

        Const ALL As String = ID & "," & SOCIO & "," & PLAN & "," & ANIO & "," & PERIODO & "," & PERIODICIDAD & "," & FECHA_VENCIMIENTO & "," &
                              FECHA_PAGO & "," & OBSERVACIONES & "," & ESTADO & "," & COBRADOR & "," & MONTO & "," & DELETED & "," & MODIFICADO & "," & BUSQUEDA_VENCIMIENTO & "," & BUSQUEDA_PAGO & "," &
                              OPERACION & "," & RECIBO_ID & "," & MOVIMIENTO_CC
    End Structure

    ''' <summary>
    ''' Tabla de cobradores
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure TABLA_COBRADORES
        Const TABLA_NOMBRE As String = "cobrador"
        Const ID As String = "cobrador_id"
        Const USER_ID As String = "cobrador_uid"
        Const NOMBRE As String = "cobrador_nombre"
        Const APELLIDO As String = "cobrador_apellido"
        Const TELEFONO As String = "cobrador_telefono"
        Const MAIL As String = "cobrador_mail"
        Const DOMICILIO As String = "cobrador_domicilio"
        Const DNI As String = "cobrador_dni"
        Const ZONA As String = "cobrador_zona"
        Const COMISION As String = "cobrador_comision"
        Const COMISION_FIJA As String = "cobrador_monto_fijo"
        Const DELETED As String = "cobrador_deleted"
        Const MODIFICADO As String = "cobrador_modificado"
        Const ALL As String = ID & "," & USER_ID & "," & NOMBRE & "," & APELLIDO & "," & TELEFONO & "," & MAIL & "," & DOMICILIO & "," & DNI & "," &
                              ZONA & "," & COMISION & "," & COMISION_FIJA & "," & DELETED & "," & MODIFICADO
    End Structure

    ''' <summary>
    ''' Tabla de usuarios del sistema
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure TABLA_USERS
        Const TABLA_NOMBRE As String = "users"
        Const ID As String = "users_id"
        Const USERNAME As String = "users_username"
        Const PASSWORD As String = "users_password"
        Const PERMISSIONS As String = "users_permissions"
        Const ADMIN As String = "users_nimda"
        Const NOMBRE As String = "users_nombre"
        Const APELLIDO As String = "users_apellido"
        Const MAIL As String = "users_mail"
        Const PROFILE As String = "users_profile"
        Const AVATAR As String = "users_avatar"
        Const TOKEN As String = "users_token"
        Const LAST_LOGIN As String = "users_lastLogin"
        Const DELETED As String = "users_deleted"
        Const MODIFICADO As String = "users_modificado"
        Const ROL As String = "usu_rol"
        Const ACTIVO As String = "usu_activo"

        Const ALL As String = ID & "," & USERNAME & "," & PASSWORD & "," & PERMISSIONS & "," & ADMIN & "," & NOMBRE & "," & APELLIDO & "," &
                              MAIL & "," & PROFILE & "," & AVATAR & "," & TOKEN & "," & LAST_LOGIN & "," & DELETED & "," & MODIFICADO
        Const ALL_TURNOS As String = ID & "," & USERNAME & "," & PASSWORD & "," & PERMISSIONS & "," & ADMIN & "," & NOMBRE & "," & APELLIDO & "," &
                              MAIL & "," & PROFILE & "," & TOKEN & "," & LAST_LOGIN & "," & DELETED & "," & MODIFICADO
    End Structure

    ''' <summary>
    ''' Tabla de permisos de usuarios del sistema
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure TABLA_PERMISOS
        Const TABLA_NOMBRE As String = "permissions"
        Const ID As String = "permissions_id"
        Const NOMBRE As String = "permissions_nombre"
        Const PERMISOS As String = "permissions_permisos"
        Const DELETED As String = "permissions_deleted"
        Const MODIFICADO As String = "permissions_modificado"
        Const ALL As String = ID & "," & NOMBRE & "," & PERMISOS & "," & DELETED & "," & MODIFICADO
    End Structure

    ''' <summary>
    ''' Tabla conteniendo las plantillas de mail
    ''' </summary>
    ''' <remarks>ALL se utiliza para evitar usar * al hacer las consultas y en el caso de agregar 
    ''' mas columnas no rompa el sistema</remarks>
    Public Structure TABLA_MAIL
        Const TABLA_NOMBRE As String = "mail"
        Const ID As String = "mail_id"
        Const NOMBRE As String = "mail_nombre"
        Const ES_HTML As String = "mail_esHTML"
        Const BODY As String = "mail_body"
        Const CSS As String = "mail_css"
        Const ADJUNTO As String = "mail_adjunto"
        Const CONTEXTO As String = "mail_contexto"
        Const DELETED As String = "mail_deleted"
        Const MODIFICADO As String = "mail_modificado"
        Const ALL As String = ID & "," & NOMBRE & "," & ES_HTML & "," & BODY & "," & CSS & "," & ADJUNTO & "," & CONTEXTO & "," & DELETED & "," & MODIFICADO
    End Structure

    Public Structure TABLA_PRONET_TARJETAS
        Const TABLA_NOMBRE As String = "Card"
        Const ID As String = TABLA_NOMBRE & "." & "ID"
        Const NOMBRE As String = TABLA_NOMBRE & "." & "Name"
        Const DESCRIPCION As String = TABLA_NOMBRE & "." & "Descr"
        Const TECHNO As String = TABLA_NOMBRE & "." & "techno"
        Const CODIGO As String = TABLA_NOMBRE & "." & "CODE"
        Const ESTADO As String = TABLA_NOMBRE & "." & "Status"
        Const DUENIO As String = TABLA_NOMBRE & "." & "Owner"
        Const S_O_C As String = TABLA_NOMBRE & "." & "SOC"
        Const TEMPLATE_ID As String = TABLA_NOMBRE & "." & "TemplateID"

        Const ALL As String = ID & "," & NOMBRE & "," & DESCRIPCION & "," & TECHNO & "," & CODIGO & "," & ESTADO & "," & DUENIO & "," & S_O_C & "," & TEMPLATE_ID
    End Structure

    Public Structure TABLA_PRONET_USUARIOS
        Const TABLA_NOMBRE As String = "CRDHLD"
        Const ID As String = TABLA_NOMBRE & "." & "ID"
        Const NUM_BADGE As String = TABLA_NOMBRE & "." & "Num_Badge"
        Const DEF As String = TABLA_NOMBRE & "." & "Default"
        Const TYPE As String = TABLA_NOMBRE & "." & "type"
        Const APELLIDO As String = TABLA_NOMBRE & "." & "Last_Name"
        Const NOMBRE As String = TABLA_NOMBRE & "." & "First_Name"
        Const NUM As String = TABLA_NOMBRE & "." & "Num"
        Const COMPANIA As String = TABLA_NOMBRE & "." & "Company"
        Const AREA As String = TABLA_NOMBRE & "." & "Area"
        Const DEPARTAMENTO As String = TABLA_NOMBRE & "." & "Department"
        Const OFFICE_PHONE As String = TABLA_NOMBRE & "." & "Office_Phone"
        Const ACCESS_GROUP As String = TABLA_NOMBRE & "." & "AccGrp"
        Const PERS_WP As String = TABLA_NOMBRE & "." & "Pers_WP"
        Const PERS_CL As String = TABLA_NOMBRE & "." & "Pers_CL"
        Const PIN As String = TABLA_NOMBRE & "." & "Pin"
        Const FROM_DATE As String = TABLA_NOMBRE & "." & "From_Date"
        Const TO_DATE As String = TABLA_NOMBRE & "." & "TO_Date"
        Const VALID As String = TABLA_NOMBRE & "." & "Valid"
        Const CARD1 As String = TABLA_NOMBRE & "." & "Card1"
        Const CARD2 As String = TABLA_NOMBRE & "." & "Card2"
        Const CARD3 As String = TABLA_NOMBRE & "." & "Card3"
    End Structure

    Public Structure TABLA_IMPORT_USUARIOS
        Const TABLA_NOMBRE As String = "importUsuarios"
        Const ID As String = "importUsuarios_id"
        Const MATRICULA As String = "importUsuarios_matricula"
        Const NOMBRE As String = "importUsuarios_nombre"
        Const APELLIDO As String = "importUsuarios_apellido"
        Const DIRECCION As String = "importUsuarios_direccion"
        Const DNI As String = "importUsuarios_dni"
        Const FIRMA As String = "importUsuarios_firma"
        Const CUIT As String = "importUsuarios_cuit"
        Const ALL As String = ID & "," & MATRICULA & "," & NOMBRE & "," & APELLIDO & "," & DIRECCION & "," & DNI & "," & FIRMA & "," & CUIT
    End Structure


    Public Structure TABLA_CONTRATOS_COFRES
        Const TABLA_NOMBRE As String = "contratoCofres"
        Const ID As String = "contratoCofres_id"
        Const TIPO As String = "contratoCofres_tipo"
        Const NUMERO As String = "contratoCofres_numero"
        Const ES_SOCIO_ID As String = "contratoCofres_esSocioId"
        Const COFRE_LETRA As String = "contratoCofres_cofreLetra"
        Const COFRE_NUMERO As String = "contratoCofres_cofreNumero"
        Const COFRE_TIPO As String = "contratoCofres_cofreTipo"
        Const NOMBRE As String = "contratoCofres_nombre"
        Const MODALIDAD_USO As String = "contratoCofres_modalidad"
        Const MOD_CONJUNTA1 As String = "contratoCofres_conjunta1"
        Const MOD_CONJUNTA2 As String = "contratoCofres_conjunta2"
        Const MOD_CONJUNTA3 As String = "contratoCofres_conjunta3"
        Const FECHA_CONTRATACION As String = "contratoCofres_fechaContratacion"
        Const FECHA_VENCIMIENTO As String = "contratoCofres_fechaVencimiento"
        Const ESTADO As String = "contratoCofres_estado"
        Const RECIBIR_INFO As String = "contratoCofres_recibirInfo"
        Const CONTACTO_CALLE As String = "contratoCofres_contactoCalle"
        Const CONTACTO_CALLE_NUM As String = "contratoCofres_contactoCalleNum"
        Const CONTACTO_CALLE_PISO As String = "contratoCofres_contactoCallePiso"
        Const CONTACTO_CALLE_DEPTO As String = "contratoCofres_contactoCalleDepto"
        Const CONTACTO_CP As String = "contratoCofres_contactoCP"
        Const CONTACTO_CIUDAD As String = "contratoCofres_contactoCiudad"
        Const CONTACTO_PROVINCIA As String = "contratoCofres_contactoProvincia"
        Const CONTACTO_TEL As String = "contratoCofres_contactoTel"
        Const CONTACTO_CEL As String = "contratoCofres_contactoCel"
        Const CONTACTO_MAIL As String = "contratoCofres_contactoMail"
        Const DELETED As String = "contratoCofres_deleted"
        Const MODIFICADO As String = "contratoCofres_modificado"
        Const ALL As String = ID & "," & TIPO & "," &
        NUMERO & "," &
        ES_SOCIO_ID & "," &
        COFRE_LETRA & "," &
        COFRE_NUMERO & "," &
        COFRE_TIPO & "," &
        NOMBRE & "," &
        MODALIDAD_USO & "," &
        MOD_CONJUNTA1 & "," &
        MOD_CONJUNTA2 & "," &
        MOD_CONJUNTA3 & "," &
        FECHA_CONTRATACION & "," &
        FECHA_VENCIMIENTO & "," &
        ESTADO & "," &
        RECIBIR_INFO & "," &
        CONTACTO_CALLE & "," &
        CONTACTO_CALLE_NUM & "," &
        CONTACTO_CALLE_PISO & "," &
        CONTACTO_CALLE_DEPTO & "," &
        CONTACTO_CP & "," &
        CONTACTO_CIUDAD & "," &
        CONTACTO_PROVINCIA & "," &
        CONTACTO_TEL & "," &
        CONTACTO_CEL & "," &
        CONTACTO_MAIL & "," &
        DELETED & "," &
        MODIFICADO
    End Structure

    Public Structure TABLA_CONTRATOS_COFRES_USUARIOS
        Const TABLA_NOMBRE As String = "contratoCofresUsuario"
        Const ID As String = "contratoCofresUsuario_id"
        Const CONTRATO_ID As String = "contratoCofresUsuario_contrato"
        Const TIPO As String = "contratoCofresUsuario_tipo"
        Const FACTURA As String = "contratoCofresUsuario_factura"
        Const IVA As String = "contratoCofresUsuario_iva"
        Const NOMBRE As String = "contratoCofresUsuario_nombre"
        Const TIPO_DNI As String = "contratoCofresUsuario_tipoDni"
        Const DNI As String = "contratoCofresUsuario_dni"
        Const CUIT As String = "contratoCofresUsuario_cuit"
        Const CONDICION_FISCAL As String = "contratoCofresUsuario_condicionFiscal"
        Const CALLE As String = "contratoCofresUsuario_calle"
        Const CALLE_NUMERO As String = "contratoCofresUsuario_calleNum"
        Const CALLE_PISO As String = "contratoCofresUsuario_callePiso"
        Const CALLE_DEPTO As String = "contratoCofresUsuario_calleDepto"
        Const CP As String = "contratoCofresUsuario_cp"
        Const CIUDAD As String = "contratoCofresUsuario_ciudad"
        Const PROVINCIA As String = "contratoCofresUsuario_provincia"
        Const TELEFONO As String = "contratoCofresUsuario_telefono"
        Const CELULAR As String = "contratoCofresUsuario_celular"
        Const MAIL As String = "contratoCofresUsuario_mail"
        Const ID_TARJETA As String = "contratoCofresUsuario_idTarjeta"
        Const FOTO As String = "contratoCofresUsuario_foto"
        Const FIRMA As String = "contratoCofresUsuario_firma"
        Const DELETED As String = "contratoCofresUsuario_deleted"
        Const MODIFICADO As String = "contratoCofresUsuario_modificado"
        Const ALL As String = ID & "," & CONTRATO_ID & "," &
        TIPO & "," &
        FACTURA & "," &
        IVA & "," &
        NOMBRE & "," &
        TIPO_DNI & "," &
        DNI & "," &
        CUIT & "," &
        CONDICION_FISCAL & "," &
        CALLE & "," &
        CALLE_NUMERO & "," &
        CALLE_PISO & "," &
        CALLE_DEPTO & "," &
        CP & "," &
        CIUDAD & "," &
        PROVINCIA & "," &
        TELEFONO & "," &
        CELULAR & "," &
        MAIL & "," &
        ID_TARJETA & "," &
        FOTO & "," &
        FIRMA & "," &
        DELETED & "," &
        MODIFICADO
    End Structure

    Public Structure TABLA_RECIBOS_DBF
        Dim SOCIO As Integer
        Dim ANIO As Integer
        Dim PERIODO As Byte
        Dim PERIODICIDAD As Byte
        Dim FECHA_VENCIMIENTO As Date
        Dim FECHA_PAGO As Date
        Dim OBSERVACIONES As String
        Dim ESTADO As Byte
        Dim SECTOR As Byte
        Dim MONTO As Double
    End Structure


    Public Structure TABLA_CATEGORIAS_IMPUESTOS
        Const TABLA_NOMBRE As String = "catImpuesto"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const NOMBRE As String = TABLA_NOMBRE & "_nombre"
        Const COMISION As String = TABLA_NOMBRE & "_comision"
        Const RETENCION As String = TABLA_NOMBRE & "_retencion"
        Const NO_COBRAR_HASTA_FECHA As String = TABLA_NOMBRE & "_noCobrarHastaFecha"
        Const FECHA_MAXIMA As String = TABLA_NOMBRE & "_fechaMaxima"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & NOMBRE & ", " & COMISION & ", " & RETENCION & ", " & NO_COBRAR_HASTA_FECHA & ", " & FECHA_MAXIMA & ", " & DELETED & ", " & MODIFICADO
    End Structure


    Public Structure TABLA_IMPUESTO
        Const TABLA_NOMBRE As String = "impuesto"
        Const ID As String = "impuesto_id"
        Const CATEGORIA As String = "impuesto_categoria"
        Const NOMBRE As String = "impuesto_nombre"
        Const USE_BARCODE As String = "impuesto_useBarcode"
        Const MSG_PROMPT As String = "impuesto_msgPrompt"
        Const BARCODE_UID As String = "impuesto_barcodeUID"
        Const BARCODE_LENGTH As String = "impuesto_barcodeLength"
        Const BARCODE_ID_SERVICIO_START As String = "impuesto_barcodeIdServicioStart"
        Const BARCODE_ID_SERVICIO_LENGTH As String = "impuesto_barcodeIdServicioLength"
        Const BARCODE_DATE_START As String = "impuesto_barcodeDateStart"
        Const BARCODE_DATE_LENGTH As String = "impuesto_barcodeDateLength"
        Const BARCODE_NUMFACTURA_START As String = "impuesto_barcodeNumFacturaStart"
        Const BARCODE_NUMFACTURA_LENGTH As String = "impuesto_barcodeNumFacturaLength"
        Const BARCODE_IMPORTE_START As String = "impuesto_barcodeImporteStart"
        Const BARCODE_IMPORTE_LENGTH As String = "impuesto_barcodeImporteLength"
        Const DELETED As String = "impuesto_deleted"
        Const MODIFICADO As String = "impuesto_modificado"
        Const ALL As String = ID & "," & CATEGORIA & "," & NOMBRE & "," & USE_BARCODE & "," & MSG_PROMPT & "," & BARCODE_UID & "," & BARCODE_LENGTH & "," & BARCODE_ID_SERVICIO_START &
                              "," & BARCODE_ID_SERVICIO_LENGTH & "," & BARCODE_DATE_START & "," & BARCODE_DATE_LENGTH & "," & BARCODE_NUMFACTURA_START & "," & BARCODE_NUMFACTURA_LENGTH & "," &
                              BARCODE_IMPORTE_START & "," & BARCODE_IMPORTE_LENGTH & "," & DELETED & "," & MODIFICADO
    End Structure

    Public Structure TABLA_PAGO_IMPUESTO
        Const TABLA_NOMBRE As String = "pagoImpuesto"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const SERVICIO_ID As String = TABLA_NOMBRE & "_servicioId"
        Const CAJERO_ID As String = TABLA_NOMBRE & "_cajeroId"
        Const CAJA As String = TABLA_NOMBRE & "_caja"
        Const IMPORTE As String = TABLA_NOMBRE & "_importe"
        Const FECHA_PAGO As String = TABLA_NOMBRE & "_fechaPago"
        Const BARCODE As String = TABLA_NOMBRE & "_barcode"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & SERVICIO_ID & ", " & CAJERO_ID & ", " & IMPORTE & ", " & FECHA_PAGO & ", " & BARCODE & ", " & DELETED & ", " & MODIFICADO
    End Structure

    Public Structure TABLA_MICROCREDITOS
        Const TABLA_NOMBRE As String = "microcreditos"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const MODO As String = TABLA_NOMBRE & "_modo"
        Const TIPO_BENEFICIARIO As String = TABLA_NOMBRE & "_tipoSocio"
        Const SOCIO As String = TABLA_NOMBRE & "_socio"
        Const MONTO As String = TABLA_NOMBRE & "_monto"
        Const GASTOS_ADMINISTRATIVOS As String = TABLA_NOMBRE & "_gastosAdministrativos"
        Const CUOTAS As String = TABLA_NOMBRE & "_cuotas"
        Const INTERES As String = TABLA_NOMBRE & "_interes"
        Const PORCENTAJE_MORA As String = TABLA_NOMBRE & "_porcentajeMora"
        Const PERIODICIDAD As String = TABLA_NOMBRE & "_periodicidad"
        Const PRIMER_VENCIMIENTO As String = TABLA_NOMBRE & "_primerVencimiento"
        Const ESTADO As String = TABLA_NOMBRE & "_estado"
        Const MONTO_ULTIMA_CUOTA As String = TABLA_NOMBRE & "_montoCuotaBonificada"
        Const TOTAL_A_PAGAR As String = TABLA_NOMBRE & "_totalPagar"
        Const CHEQUE As String = TABLA_NOMBRE & "_cheque"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & MODO & ", " & TIPO_BENEFICIARIO & ", " & SOCIO & ", " & MONTO & ", " & GASTOS_ADMINISTRATIVOS & ", " & CUOTAS & ", " & INTERES _
                               & ", " & PORCENTAJE_MORA & ", " & PERIODICIDAD & ", " & PRIMER_VENCIMIENTO & ", " & ESTADO _
                                & ", " & MONTO_ULTIMA_CUOTA & ", " & TOTAL_A_PAGAR & ", " & CHEQUE & ", " & DELETED & ", " & MODIFICADO
    End Structure

    Public Structure TABLA_CUOTAS_MICROCREDITOS
        Const TABLA_NOMBRE As String = "mcCuotas"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const MICROCREDITO As String = TABLA_NOMBRE & "_microcredito"
        Const NUMERO As String = TABLA_NOMBRE & "_numero"
        Const IMPORTE As String = TABLA_NOMBRE & "_importe"
        Const FECHA_VENCIMIENTO As String = TABLA_NOMBRE & "_fechaVencimiento"
        Const DETALLE As String = TABLA_NOMBRE & "_detalle"
        Const ESTADO As String = TABLA_NOMBRE & "_estado"
        Const FECHA_PAGO As String = TABLA_NOMBRE & "_fechaPago"
        Const EN_MORA As String = TABLA_NOMBRE & "_enMora"
        Const INTERESES As String = TABLA_NOMBRE & "_intereses"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const OPERACION As String = TABLA_NOMBRE & "_operacion"
        Const ALL As String = ID & ", " & MICROCREDITO & ", " & NUMERO & ", " & IMPORTE & ", " & FECHA_VENCIMIENTO & ", " & DETALLE & ", " & ESTADO & ", " & FECHA_PAGO & ", " & EN_MORA & ", " & INTERESES & ", " & DELETED & ", " & MODIFICADO & ", " & OPERACION
    End Structure

    Public Structure TABLA_PERSONAL
        Const TABLA_NOMBRE As String = "personal"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const NOMBRE As String = TABLA_NOMBRE & "_nombre"
        Const APELLIDO As String = TABLA_NOMBRE & "_apellido"
        Const FUNCIONES As String = TABLA_NOMBRE & "_funciones"
        Const DOMICILIO As String = TABLA_NOMBRE & "_domicilio"
        Const MAIL As String = TABLA_NOMBRE & "_mail"
        Const TELEFONO As String = TABLA_NOMBRE & "_telefono"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & NOMBRE & ", " & APELLIDO & ", " & FUNCIONES & ", " &
                              DOMICILIO & ", " & MAIL & ", " & TELEFONO & ", " & DELETED & ", " & MODIFICADO
    End Structure

    Public Structure TABLA_PAGOS_COFRES
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
        Const OPERACION As String = TABLA_NOMBRE & "_operacion"
        Const ALL As String = ID & ", " & CONTRATO & ", " & IMPORTE & ", " & PERIODO & ", " & ANIO & ", " & FECHA_PAGO & ", " & ESTADO & ", " & DELETED & ", " & MODIFICADO & ", " & OPERACION
    End Structure

    Public Structure TABLA_RECIBOS
        Const TABLA_NOMBRE As String = "recibos"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const PUNTO_VENTA As String = TABLA_NOMBRE & "_puntoVenta"
        Const NUMERO As String = TABLA_NOMBRE & "_numero"
        Const FECHA As String = TABLA_NOMBRE & "_fecha"
        Const CLIENTE_TIPO As String = TABLA_NOMBRE & "_clienteTipo"
        Const CLIENTE_ID As String = TABLA_NOMBRE & "_clienteId"
        Const CONCEPTO As String = TABLA_NOMBRE & "_concepto"
        Const IMPORTE As String = TABLA_NOMBRE & "_importe"
        Const EMISOR As String = TABLA_NOMBRE & "_emisor"
        Const HORA_EMISION As String = TABLA_NOMBRE & "_horaEmision"
        Const CAJA As String = TABLA_NOMBRE & "_caja"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const OPERACION As String = TABLA_NOMBRE & "_operacion"
        Const ALL As String = ID & ", " & PUNTO_VENTA & ", " & NUMERO & ", " & FECHA & ", " & CLIENTE_TIPO & ", " & CLIENTE_ID & ", " & CONCEPTO & ", " & IMPORTE & ", " & EMISOR & ", " & HORA_EMISION & ", " & CAJA & ", " & DELETED & ", " & MODIFICADO & ", " & OPERACION
    End Structure

    Public Structure TABLA_FACTURA
        Const TABLA_NOMBRE As String = "factura"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const PUNTO_VENTA As String = TABLA_NOMBRE & "_puntoVenta"
        Const NUMERO As String = TABLA_NOMBRE & "_numero"
        Const FECHA As String = TABLA_NOMBRE & "_fecha"
        Const CLIENTE_TIPO As String = TABLA_NOMBRE & "_clienteTipo"
        Const CLIENTE_ID As String = TABLA_NOMBRE & "_clienteId"
        Const CONCEPTO As String = TABLA_NOMBRE & "_concepto"
        Const IMPORTE As String = TABLA_NOMBRE & "_importe"
        Const EMISOR As String = TABLA_NOMBRE & "_emisor"
        Const HORA_EMISION As String = TABLA_NOMBRE & "_horaEmision"
        Const CAJA As String = TABLA_NOMBRE & "_caja"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const OPERACION As String = TABLA_NOMBRE & "_operacion"
        Const ALL As String = ID & ", " & PUNTO_VENTA & ", " & NUMERO & ", " & FECHA & ", " & CLIENTE_TIPO & ", " & CLIENTE_ID & ", " & CONCEPTO & ", " & IMPORTE & ", " & EMISOR & ", " & HORA_EMISION & ", " & CAJA & ", " & DELETED & ", " & MODIFICADO & ", " & OPERACION
    End Structure

    Public Structure TABLA_FACTURA_DETALLE
        Const TABLA_NOMBRE As String = "facturaDetalle"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const FACTURA_ID As String = TABLA_NOMBRE & "_facturaId"
        Const CANTIDAD As String = TABLA_NOMBRE & "_cantidad"
        Const DETALLE As String = TABLA_NOMBRE & "_detalle"
        Const P_UNITARIO As String = TABLA_NOMBRE & "_pUnitario"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & FACTURA_ID & ", " & CANTIDAD & ", " & DETALLE & ", " & P_UNITARIO & ", " & DELETED & ", " & MODIFICADO
    End Structure

    Public Structure TABLA_BLOQUEO_COFRES
        Const TABLA_NOMBRE As String = "bloqueoCofres"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const CONTRATO_ID As String = TABLA_NOMBRE & "_contratoId"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & CONTRATO_ID & ", " & DELETED & ", " & MODIFICADO
    End Structure

    Public Structure TABLA_BLOQUEO_TARJETAS
        Const TABLA_NOMBRE As String = "bloqueoTarjetas"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const BLOQUEO_COFRE_ID As String = TABLA_NOMBRE & "_bloqueoCofreId"
        Const PRONET_CARD_HOLDER_ID As String = TABLA_NOMBRE & "_pronetCardHolderId"
        Const PRONET_CARD_HOLDER_CODE As String = TABLA_NOMBRE & "_pronetCardHolderCode"
        Const PRONET_CARD_HOLDER_STATUS As String = TABLA_NOMBRE & "_pronetCardHolderStatus"
        Const PRONET_CARD_HOLDER_OWNER As String = TABLA_NOMBRE & "_pronetCardHolderOwner"
        Const PRONET_CARD_ID As String = TABLA_NOMBRE & "_pronetCardId"
        Const PRONET_CARD_STATUS As String = TABLA_NOMBRE & "_pronetCardStatus"
        Const PRONET_CARD_OWNER As String = TABLA_NOMBRE & "_pronetCardOwner"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & BLOQUEO_COFRE_ID & ", " & PRONET_CARD_HOLDER_ID & ", " & PRONET_CARD_HOLDER_CODE & ", " & PRONET_CARD_HOLDER_STATUS & ", " & PRONET_CARD_HOLDER_OWNER & ", " & PRONET_CARD_ID & ", " & PRONET_CARD_STATUS & ", " & PRONET_CARD_OWNER & ", " & DELETED & ", " & MODIFICADO
    End Structure

    Public Structure TABLA_PRONET_JOURNAL
        Const TABLA_NOMBRE As String = "LOG"
        Const ID As String = "ID"
        Const FECHA As String = "DATE"
        Const LOCACION As String = "From_Name"
        Const USUARIO As String = "Desc3"
        Const PUERTA As String = "Reader"
        Const TARJETA As String = "Cardholder"
        Const ALL = ID & ", " & FECHA & ", " & LOCACION & ", " & USUARIO & ", " & PUERTA & ", " & TARJETA
    End Structure

    Public Structure TABLA_EVENTOS_MAIL
        Const TABLA_NOMBRE As String = "eventosMail"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const NOMBRE As String = TABLA_NOMBRE & "_nombre"
        Const PLANTILLA As String = TABLA_NOMBRE & "_plantilla"
        Const TIPO_FECHA As String = TABLA_NOMBRE & "_tipoFecha"
        Const FECHA_FIJA As String = TABLA_NOMBRE & "_fechaFija"
        Const FECHA_DIA_SEMANA As String = TABLA_NOMBRE & "_fechaDiaSemana"
        Const FECHA_MES As String = TABLA_NOMBRE & "_fechaMes"
        Const DIAS_RETRASO As String = TABLA_NOMBRE & "_diasRetraso"
        Const DESTINATARIOS As String = TABLA_NOMBRE & "_destinatarios"
        Const ELIMINAR_AL_FINALIZAR As String = TABLA_NOMBRE & "_eliminarAlFinalizar"
        Const ULTIMA_EJECUCION As String = TABLA_NOMBRE & "_ultimaEjecucion"
        Const ASUNTO As String = TABLA_NOMBRE & "_asunto"
        Const ADJUNTO As String = TABLA_NOMBRE & "_adjunto"
        Const ULTIMO_I_D As String = TABLA_NOMBRE & "_ultimoID"
        Const ULTIMO_D_T As String = TABLA_NOMBRE & "_ultimoDT"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & NOMBRE & ", " & PLANTILLA & ", " & TIPO_FECHA & ", " & FECHA_FIJA & ", " & FECHA_DIA_SEMANA & ", " & FECHA_MES & ", " & DIAS_RETRASO & ", " & DESTINATARIOS & ", " & ELIMINAR_AL_FINALIZAR & ", " & ULTIMA_EJECUCION & ", " & ASUNTO & ", " & ADJUNTO & ", " & ULTIMO_I_D & ", " & ULTIMO_D_T & ", " & DELETED & ", " & MODIFICADO
    End Structure

    Public Structure TABLA_LOG
        Const TABLA_NOMBRE As String = "log"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const EVENT_I_D As String = TABLA_NOMBRE & "_eventID"
        Const LEVEL As String = TABLA_NOMBRE & "_level"
        Const SOURCE As String = TABLA_NOMBRE & "_source"
        Const DESCRIPTION As String = TABLA_NOMBRE & "_description"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & EVENT_I_D & ", " & LEVEL & ", " & SOURCE & ", " & DESCRIPTION & ", " & DELETED & ", " & MODIFICADO
    End Structure

    Public Structure TABLA_SETTINGS
        Const TABLA_NOMBRE As String = "settings"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const MAIL_USERNAME As String = TABLA_NOMBRE & "_mailUsername"
        Const MAIL_PASSWORD As String = TABLA_NOMBRE & "_mailPassword"
        Const MAIL_HOST As String = TABLA_NOMBRE & "_mailHost"
        Const MAIL_PORT As String = TABLA_NOMBRE & "_mailPort"
        Const MAIL_SSL As String = TABLA_NOMBRE & "_mailSsl"
        Const MAIL_FROM_NAME As String = TABLA_NOMBRE & "_mailFromName"
        Const MAIL_GMAIL_USERNAME As String = TABLA_NOMBRE & "_mailGmailUsername"
        Const MAIL_GMAIL_PASSWORD As String = TABLA_NOMBRE & "_mailGmailPassword"
        Const MAIL_GMAIL_SMTP As String = TABLA_NOMBRE & "_mailGmailSmtp"
        Const MAIL_GMAIL_SMTP_PORT As String = TABLA_NOMBRE & "_mailGmailSmtpPort"
        Const MAIL_GMAIL_SMTP_SSL As String = TABLA_NOMBRE & "_mailGmailSmtpSsl"
        Const MAIL_TEST_ADDRESS As String = TABLA_NOMBRE & "_mailTestAddress"
        Const MAIL_RESPONSE_ADDRESS As String = TABLA_NOMBRE & "_mailResponseAddress"
        Const SYSTEM_PRONET_PATH As String = TABLA_NOMBRE & "_systemPronetPath"
        Const FACTURA_OFFSET_X As String = TABLA_NOMBRE & "_facturaOffsetX"
        Const FACTURA_OFFSET_Y As String = TABLA_NOMBRE & "_facturaOffsetY"
        Const RECIBO_OFFSET_X As String = TABLA_NOMBRE & "_reciboOffsetX"
        Const RECIBO_OFFSET_Y As String = TABLA_NOMBRE & "_reciboOffsetY"
        Const SOCIOS_OFFSET_X As String = TABLA_NOMBRE & "_sociosOffsetX"
        Const SOCIOS_OFFSET_Y As String = TABLA_NOMBRE & "_sociosOffsetY"
        Const SYSTEM_BARCODE_FONT As String = TABLA_NOMBRE & "_systemBarcodeFont"
        Const SYSTEM_DEBUG_MODE As String = TABLA_NOMBRE & "_systemDebugMode"
        Const SYSTEM_BACKUP_DIR As String = TABLA_NOMBRE & "_systemBackupDir"
        Const SYSTEM_BACKUP_HOUR As String = TABLA_NOMBRE & "_systemBackupHour"
        Const SYSTEM_BACKUP_RETRY_MINUTES As String = TABLA_NOMBRE & "_systemBackupRetryMinutes"
        Const SYSTEM_BACKUP_TOTAL_RETRY As String = TABLA_NOMBRE & "_systemBackupTotalRetry"
        Const TURNOS_DB_PATH As String = TABLA_NOMBRE & "_turnosDbPath"
        Const TURNOS_DB_NAME As String = TABLA_NOMBRE & "_turnosDbName"
        Const TURNOS_DB_USER As String = TABLA_NOMBRE & "_turnosDbUser"
        Const TURNOS_DB_PASSWORD As String = TABLA_NOMBRE & "_turnosDbPassword"
        Const TURNOS_DB_PORT As String = TABLA_NOMBRE & "_turnosDbPort"
        Const TURNOS_DB_URL As String = TABLA_NOMBRE & "_turnosDbUrl"
        Const ALL As String = ID & ", " & MAIL_USERNAME & ", " & MAIL_PASSWORD & ", " & MAIL_HOST & ", " & MAIL_PORT & ", " & MAIL_SSL & ", " & MAIL_FROM_NAME & ", " & MAIL_GMAIL_USERNAME & ", " & MAIL_GMAIL_PASSWORD & ", " & MAIL_GMAIL_SMTP & ", " & MAIL_GMAIL_SMTP_PORT & ", " & MAIL_GMAIL_SMTP_SSL & ", " & MAIL_TEST_ADDRESS & ", " & MAIL_RESPONSE_ADDRESS & ", " & SYSTEM_PRONET_PATH & ", " & FACTURA_OFFSET_X & ", " & FACTURA_OFFSET_Y & ", " & RECIBO_OFFSET_X & ", " & RECIBO_OFFSET_Y & ", " & SOCIOS_OFFSET_X & ", " & SOCIOS_OFFSET_Y & ", " & SYSTEM_BARCODE_FONT & ", " & SYSTEM_DEBUG_MODE & ", " & SYSTEM_BACKUP_DIR & ", " & SYSTEM_BACKUP_HOUR & ", " & SYSTEM_BACKUP_RETRY_MINUTES & ", " & SYSTEM_BACKUP_TOTAL_RETRY & ", " & TURNOS_DB_PATH & ", " & TURNOS_DB_NAME & ", " & TURNOS_DB_USER & ", " & TURNOS_DB_PASSWORD & ", " & TURNOS_DB_PORT & ", " & TURNOS_DB_URL
    End Structure

    Public Structure TABLA_SOCIOS_RUBRO
        Const TABLA_NOMBRE As String = "sociosRubro"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const NOMBRE As String = TABLA_NOMBRE & "_nombre"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & NOMBRE & ", " & DELETED & ", " & MODIFICADO
    End Structure

    Public Structure TABLA_SOCIOS_SEGMENTO
        Const TABLA_NOMBRE As String = "sociosSegmento"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const RUBRO As String = TABLA_NOMBRE & "_rubro"
        Const NOMBRE As String = TABLA_NOMBRE & "_nombre"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & RUBRO & ", " & NOMBRE & ", " & DELETED & ", " & MODIFICADO
    End Structure

    Public Structure TABLA_LOCALIDADES
        Const TABLA_NOMBRE As String = "Localidad"
        Const ID As String = "Localidad.ID"
        Const ID_DEPARTAMENTO As String = "Localidad.idDepartamento"
        Const NOMBRE As String = "Localidad.Nombre"
        Const ALL As String = ID & ", " & ID_DEPARTAMENTO & ", " & NOMBRE
    End Structure


End Module
