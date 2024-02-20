Imports SelectPdf
Imports QRCoder

Public Class AfipFERenderer
    Public Property RenderQR As Boolean = False
    Public Function CSSFull() As String
        Dim tmpStr As String = ""
        tmpStr &= "<style>" & vbCrLf
        tmpStr &= "body,html {" & vbCrLf
        tmpStr &= "padding : 0;" & vbCrLf
        tmpStr &= "margin: 0;" & vbCrLf
        tmpStr &= "font-family:Arial,sans-serif;" & vbCrLf
        tmpStr &= "line-height:1;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "p {" & vbCrLf
        tmpStr &= "margin:0;" & vbCrLf
        tmpStr &= "letter-spacing:.9;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".fe {" & vbCrLf
        tmpStr &= "position:relative;" & vbCrLf
        tmpStr &= "width:210mm;" & vbCrLf
        tmpStr &= "height:270mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".debug {" & vbCrLf
        tmpStr &= "border:1px solid red;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".content {" & vbCrLf
        tmpStr &= "position : absolute;" & vbCrLf
        tmpStr &= "display:block;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".originalDuplicado {" & vbCrLf
        tmpStr &= "width:100%;" & vbCrLf
        tmpStr &= "text-align:center;" & vbCrLf
        tmpStr &= "top:8.3mm;" & vbCrLf
        tmpStr &= "left : 0;" & vbCrLf
        tmpStr &= "font-weight: bold;" & vbCrLf
        tmpStr &= "font-size:4.6mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".tipoComprobante {" & vbCrLf
        tmpStr &= "top:15.184mm;" & vbCrLf
        tmpStr &= "left:96.672mm;" & vbCrLf
        tmpStr &= "width:16.384mm;" & vbCrLf
        tmpStr &= "height:14.817mm;" & vbCrLf
        tmpStr &= "text-align:center;" & vbCrLf
        tmpStr &= "z-index:4;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".tipoComprobante .tipo {" & vbCrLf
        tmpStr &= "margin-top:1mm;" & vbCrLf
        tmpStr &= "font-weight: bold;" & vbCrLf
        tmpStr &= "font-size:8mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".tipoComprobante .tipoCodigo {" & vbCrLf
        tmpStr &= "font-size:2.8mm;" & vbCrLf
        tmpStr &= "margin-top:.5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleComprobante, .headerComprobante {" & vbCrLf
        tmpStr &= "left:120.092mm;" & vbCrLf
        tmpStr &= "top:19.946mm;" & vbCrLf
        tmpStr &= "width:90.418mm;" & vbCrLf
        tmpStr &= "height:34.793mm;" & vbCrLf
        tmpStr &= "font-size:2.8mm;" & vbCrLf
        tmpStr &= "line-height:1;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".headerComprobante {" & vbCrLf
        tmpStr &= "left:7.447mm;" & vbCrLf
        tmpStr &= "top: 45.252mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".logo {" & vbCrLf
        tmpStr &= "left:7.447mm;" & vbCrLf
        tmpStr &= "top:21.435mm;" & vbCrLf
        tmpStr &= "height:20.362mm;" & vbCrLf
        tmpStr &= "width:87mm;" & vbCrLf
        tmpStr &= "text-align:center;" & vbCrLf
        tmpStr &= "font-size: 6.13mm;" & vbCrLf
        tmpStr &= "font-weight:bold;" & vbCrLf
        tmpStr &= "line-height:1.3;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleComprobante .tipo{" & vbCrLf
        tmpStr &= "font-size: 6.13mm;" & vbCrLf
        tmpStr &= "margin-bottom: 5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleComprobante .text, .headerComprobante .text{" & vbCrLf
        tmpStr &= "margin-bottom:.8mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".fechasComprobante{" & vbCrLf
        tmpStr &= "font-size: 3.5mm;" & vbCrLf
        tmpStr &= "left:7.447mm;" & vbCrLf
        tmpStr &= "top: 60mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleReceptor {" & vbCrLf
        tmpStr &= "font-size: 2.8mm;" & vbCrLf
        tmpStr &= "width: 75mm;" & vbCrLf
        tmpStr &= "left:7.447mm;" & vbCrLf
        tmpStr &= "top: 67.4mm;" & vbCrLf
        tmpStr &= "line-height:5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleReceptorB {" & vbCrLf
        tmpStr &= "font-size: 2.8mm;" & vbCrLf
        tmpStr &= "width: 115mm;" & vbCrLf
        tmpStr &= "left:87.447mm;" & vbCrLf
        tmpStr &= "top: 67.4mm;" & vbCrLf
        tmpStr &= "line-height:5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleHeaders {" & vbCrLf
        tmpStr &= "font-size: 2.8mm;" & vbCrLf
        tmpStr &= "left:7.447mm;" & vbCrLf
        tmpStr &= "top: 90.2mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleHeaders div {" & vbCrLf
        tmpStr &= "position: absolute;" & vbCrLf
        tmpStr &= "top:0;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleHeaders div b {" & vbCrLf
        tmpStr &= "width: 28mm;" & vbCrLf
        tmpStr &= "display:block;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleContainer {" & vbCrLf
        tmpStr &= "font-size: 2.8mm;" & vbCrLf
        tmpStr &= "left:5mm;" & vbCrLf
        tmpStr &= "top: 96.7mm;" & vbCrLf
        tmpStr &= "width:199.5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleContainer table tr {" & vbCrLf
        tmpStr &= "vertical-align:top;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleContainer table tr td{" & vbCrLf
        tmpStr &= "padding-bottom: 2mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerTotalHeader, .footerTotalDatos {" & vbCrLf
        tmpStr &= "text-align: right;" & vbCrLf
        tmpStr &= "left:133.721mm;" & vbCrLf
        tmpStr &= "top:215.4mm;" & vbCrLf
        tmpStr &= "font-size:3.2mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerTotalDatos {" & vbCrLf
        tmpStr &= "width: 25mm;" & vbCrLf
        tmpStr &= "left: 172mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerTotalHeader p, , .footerTotalDatos p {" & vbCrLf
        tmpStr &= "padding-bottom:4mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerTotalHeader .text, , .footerTotalDatos .text {" & vbCrLf
        tmpStr &= "font-size:3.6mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerBarcode{" & vbCrLf
        tmpStr &= "left:6mm;" & vbCrLf
        tmpStr &= "top:247.5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".barcodeContainer {" & vbCrLf
        tmpStr &= "height: 13mm;" & vbCrLf
        tmpStr &= "overflow:hidden;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerBarcode .barcode {" & vbCrLf
        tmpStr &= "font-family: ""Code 2 of 5 Interleaved"", sans-serif;" & vbCrLf
        tmpStr &= "text-align:left;" & vbCrLf
        tmpStr &= "font-size : 12mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerBarcode .barcodeNumber {" & vbCrLf
        tmpStr &= "font-size : 2.8mm;" & vbCrLf
        tmpStr &= "padding-left:12mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerBarcode small {" & vbCrLf
        tmpStr &= "font-weight: bold;" & vbCrLf
        tmpStr &= "font-style: italic;" & vbCrLf
        tmpStr &= "font-size: 2.1mm;" & vbCrLf
        tmpStr &= "padding-left: 2mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".comprobanteAutorizado{" & vbCrLf
        tmpStr &= "left:41mm;" & vbCrLf
        tmpStr &= "top:241mm;" & vbCrLf
        tmpStr &= "font-weight:bold;" & vbCrLf
        tmpStr &= "font-style:italic;" & vbCrLf
        tmpStr &= "font-size:3.19mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".paginacion {" & vbCrLf
        tmpStr &= "left:98mm;" & vbCrLf
        tmpStr &= "top:241mm;" & vbCrLf
        tmpStr &= "font-weight:bold;" & vbCrLf
        tmpStr &= "font-size:3.3mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".caeHeader, .caeDatos {" & vbCrLf
        tmpStr &= "font-weight: bold;" & vbCrLf
        tmpStr &= "left:128mm;" & vbCrLf
        tmpStr &= "top:239mm;" & vbCrLf
        tmpStr &= "text-align:right;" & vbCrLf
        tmpStr &= "font-size:3.3mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".caeHeader p, .caeDatos p {" & vbCrLf
        tmpStr &= "margin-bottom: 2.5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".caeDatos {" & vbCrLf
        tmpStr &= "font-weight: normal;" & vbCrLf
        tmpStr &= "left:165mm;" & vbCrLf
        tmpStr &= "text-align:left;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".blk {" & vbCrLf
        tmpStr &= "display: block;" & vbCrLf
        tmpStr &= "position: absolute;" & vbCrLf
        tmpStr &= "border: 1px solid #444;" & vbCrLf
        tmpStr &= "width: 199.6mm;" & vbCrLf
        tmpStr &= "top: 0;" & vbCrLf
        tmpStr &= "left: 5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".header-blk {" & vbCrLf
        tmpStr &= "height: 51mm;" & vbCrLf
        tmpStr &= "top:5.5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".date-blk {" & vbCrLf
        tmpStr &= "height:7.8mm;" & vbCrLf
        tmpStr &= "top: 57.5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".client-blk {" & vbCrLf
        tmpStr &= "height:21.45mm;" & vbCrLf
        tmpStr &= "top: 66mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".det-blk {" & vbCrLf
        tmpStr &= "height:5.85mm;" & vbCrLf
        tmpStr &= "top: 88.51mm;" & vbCrLf
        tmpStr &= "background:#cccccc;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".price-blk {" & vbCrLf
        tmpStr &= "height:33.50mm;" & vbCrLf
        tmpStr &= "top: 202mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".blk-inner {" & vbCrLf
        tmpStr &= "top:0;" & vbCrLf
        tmpStr &= "left:0;" & vbCrLf
        tmpStr &= "border: none;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".header-orig-blk{" & vbCrLf
        tmpStr &= "border-bottom: 1px solid #444;" & vbCrLf
        tmpStr &= "height: 9mm;" & vbCrLf
        tmpStr &= "z-index: 2;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".header-type-blk{" & vbCrLf
        tmpStr &= "border: 1px solid #444;" & vbCrLf
        tmpStr &= "width: 16.1mm;" & vbCrLf
        tmpStr &= "height: 15mm;" & vbCrLf
        tmpStr &= "border-top: none;" & vbCrLf
        tmpStr &= "left: 91.3mm;" & vbCrLf
        tmpStr &= "top:9.1mm;" & vbCrLf
        tmpStr &= "background: white;" & vbCrLf
        tmpStr &= "z-index: 1;" & vbCrLf
        tmpStr &= "}" & vbCrLf

        tmpStr &= ".header-logo-blk{" & vbCrLf
        tmpStr &= "border-right: 1px solid #444;" & vbCrLf
        tmpStr &= "width: 16.1mm;" & vbCrLf
        tmpStr &= "height: 27mm;" & vbCrLf
        tmpStr &= "left: 83.5mm;" & vbCrLf
        tmpStr &= "top:24.3mm;" & vbCrLf
        tmpStr &= "z-index: 0;" & vbCrLf
        tmpStr &= "}" & vbCrLf


        tmpStr &= ".det-codigo-blk{" & vbCrLf
        tmpStr &= "border-right: 1px solid #444;" & vbCrLf
        tmpStr &= "width: 13.8mm;" & vbCrLf
        tmpStr &= "height: 100%;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".det-producto-blk{" & vbCrLf
        tmpStr &= "border-right: 1px solid #444;" & vbCrLf
        tmpStr &= "width: 63.5mm;" & vbCrLf
        tmpStr &= "height: 100%;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".det-cantidad-blk{" & vbCrLf
        tmpStr &= "border-right: 1px solid #444;" & vbCrLf
        tmpStr &= "width: 86.3mm;" & vbCrLf
        tmpStr &= "height: 100%;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".det-medida-blk{" & vbCrLf
        tmpStr &= "border-right: 1px solid #444;" & vbCrLf
        tmpStr &= "width: 101.3mm;" & vbCrLf
        tmpStr &= "height: 100%;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".det-precio-blk{" & vbCrLf
        tmpStr &= "border-right: 1px solid #444;" & vbCrLf
        tmpStr &= "width: 129.58mm;" & vbCrLf
        tmpStr &= "height: 100%;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".det-bonif-blk{" & vbCrLf
        tmpStr &= "border-right: 1px solid #444;" & vbCrLf
        tmpStr &= "width: 141.2mm;" & vbCrLf
        tmpStr &= "height: 100%;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".det-imp-blk{" & vbCrLf
        tmpStr &= "border-right: 1px solid #444;" & vbCrLf
        tmpStr &= "width: 166.8mm;" & vbCrLf
        tmpStr &= "height: 100%;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".logo-afip{" & vbCrLf
        tmpStr &= "position: absolute;" & vbCrLf
        tmpStr &= "width: 120px;" & vbCrLf
        tmpStr &= "top: -9mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".breakafter{" & vbCrLf
        tmpStr &= "page-break-after: always;" & vbCrLf
        tmpStr &= "}" & vbCrLf



        ' Nuevo QR
        If RenderQR Then
            tmpStr &= ".qrcode {" & vbCrLf
            tmpStr &= "position: absolute;" & vbCrLf
            tmpStr &= "width: 30mm;" & vbCrLf
            tmpStr &= "top: -10mm;" & vbCrLf
            tmpStr &= "}" & vbCrLf
            tmpStr &= ".logo-afip{" & vbCrLf
            tmpStr &= "position: absolute;" & vbCrLf
            tmpStr &= "width: 120px;" & vbCrLf
            tmpStr &= "top: -9mm;" & vbCrLf
            tmpStr &= "left: 33mm;" & vbCrLf
            tmpStr &= "}" & vbCrLf
            tmpStr &= ".barcodeContainer {" & vbCrLf
            tmpStr &= "height: 13mm;" & vbCrLf
            tmpStr &= "overflow:hidden;" & vbCrLf
            tmpStr &= "}" & vbCrLf
            tmpStr &= ".footerBarcode .barcode {" & vbCrLf
            tmpStr &= "display: none;" & vbCrLf
            tmpStr &= "}" & vbCrLf
            tmpStr &= ".footerBarcode .barcodeNumber {" & vbCrLf
            tmpStr &= "display: none;" & vbCrLf
            tmpStr &= "}" & vbCrLf
            tmpStr &= ".footerBarcode small {" & vbCrLf
            tmpStr &= "display: block;" & vbCrLf
            tmpStr &= "width: 100mm;" & vbCrLf
            tmpStr &= "position: absolute;" & vbCrLf
            tmpStr &= "padding-left: 32mm;" & vbCrLf
            tmpStr &= "top: 10mm;" & vbCrLf
            tmpStr &= "}" & vbCrLf
            tmpStr &= ".comprobanteAutorizado{" & vbCrLf
            tmpStr &= "left: 38mm;" & vbCrLf
            tmpStr &= "top:250mm;" & vbCrLf
            tmpStr &= "}" & vbCrLf
        End If


        tmpStr &= "</style>" & vbCrLf
        Return tmpStr
    End Function

    Public Function CSSStyleLegacy() As String
        Dim tmpStr As String = ""
        tmpStr &= "<style>" & vbCrLf
        tmpStr &= "body,html {" & vbCrLf
        tmpStr &= "padding : 0;" & vbCrLf
        tmpStr &= "margin: 0;" & vbCrLf
        tmpStr &= "font-family:Arial,sans-serif;" & vbCrLf
        tmpStr &= "line-height:1;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "p {" & vbCrLf
        tmpStr &= "margin:0;" & vbCrLf
        tmpStr &= "letter-spacing:.9;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".fe {" & vbCrLf
        tmpStr &= "position:relative;" & vbCrLf
        tmpStr &= "width:210mm;" & vbCrLf
        tmpStr &= "height:280mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".fondo {" & vbCrLf
        tmpStr &= "width:100%;" & vbCrLf
        tmpStr &= "opacity:.3;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".debug {" & vbCrLf
        tmpStr &= "border:1px solid red;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".content {" & vbCrLf
        tmpStr &= "position : absolute;" & vbCrLf
        tmpStr &= "display:block;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".originalDuplicado {" & vbCrLf
        tmpStr &= "width:100%;" & vbCrLf
        tmpStr &= "text-align:center;" & vbCrLf
        tmpStr &= "top:8.3mm;" & vbCrLf
        tmpStr &= "left : 0;" & vbCrLf
        tmpStr &= "font-weight: bold;" & vbCrLf
        tmpStr &= "font-size:4.6mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".tipoComprobante {" & vbCrLf
        tmpStr &= "top:15.184mm;" & vbCrLf
        tmpStr &= "left:96.672mm;" & vbCrLf
        tmpStr &= "width:16.384mm;" & vbCrLf
        tmpStr &= "height:14.817mm;" & vbCrLf
        tmpStr &= "text-align:center;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".tipoComprobante .tipo {" & vbCrLf
        tmpStr &= "margin-top:1mm;" & vbCrLf
        tmpStr &= "font-weight: bold;" & vbCrLf
        tmpStr &= "font-size:8mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".tipoComprobante .tipoCodigo {" & vbCrLf
        tmpStr &= "font-size:2.8mm;" & vbCrLf
        tmpStr &= "margin-top:.5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleComprobante, .headerComprobante {" & vbCrLf
        tmpStr &= "left:120.092mm;" & vbCrLf
        tmpStr &= "top:19.946mm;" & vbCrLf
        tmpStr &= "width:82.418mm;" & vbCrLf
        tmpStr &= "height:34.793mm;" & vbCrLf
        tmpStr &= "font-size:2.8mm;" & vbCrLf
        tmpStr &= "line-height:1;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".headerComprobante {" & vbCrLf
        tmpStr &= "left:7.447mm;" & vbCrLf
        tmpStr &= "top: 45.252mm;" & vbCrLf
        tmpStr &= "width:87.418mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".logo {" & vbCrLf
        tmpStr &= "left:7.447mm;" & vbCrLf
        tmpStr &= "top:21.435mm;" & vbCrLf
        tmpStr &= "height:20.362mm;" & vbCrLf
        tmpStr &= "width:87mm;" & vbCrLf
        tmpStr &= "text-align:center;" & vbCrLf
        tmpStr &= "font-size: 6.13mm;" & vbCrLf
        tmpStr &= "font-weight:bold;" & vbCrLf
        tmpStr &= "line-height:1.3;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleComprobante .tipo{" & vbCrLf
        tmpStr &= "font-size: 6.13mm;" & vbCrLf
        tmpStr &= "margin-bottom: 5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleComprobante .text, .headerComprobante .text{" & vbCrLf
        tmpStr &= "margin-bottom:.8mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".fechasComprobante{" & vbCrLf
        tmpStr &= "font-size: 3.5mm;" & vbCrLf
        tmpStr &= "left:7.447mm;" & vbCrLf
        tmpStr &= "top: 60mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleReceptor {" & vbCrLf
        tmpStr &= "font-size: 2.8mm;" & vbCrLf
        tmpStr &= "width: 75mm;" & vbCrLf
        tmpStr &= "left:7.447mm;" & vbCrLf
        tmpStr &= "top: 67.4mm;" & vbCrLf
        tmpStr &= "line-height:5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleReceptorB {" & vbCrLf
        tmpStr &= "font-size: 2.8mm;" & vbCrLf
        tmpStr &= "width: 115mm;" & vbCrLf
        tmpStr &= "left:87.447mm;" & vbCrLf
        tmpStr &= "top: 67.4mm;" & vbCrLf
        tmpStr &= "line-height:5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleHeaders {" & vbCrLf
        tmpStr &= "font-size: 2.8mm;" & vbCrLf
        tmpStr &= "left:7.447mm;" & vbCrLf
        tmpStr &= "top: 90.2mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleHeaders div {" & vbCrLf
        tmpStr &= "position: absolute;" & vbCrLf
        tmpStr &= "top:0;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleHeaders div b {" & vbCrLf
        tmpStr &= "width: 28mm;" & vbCrLf
        tmpStr &= "display:block;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleContainer {" & vbCrLf
        tmpStr &= "font-size: 2.8mm;" & vbCrLf
        tmpStr &= "left:5mm;" & vbCrLf
        tmpStr &= "top: 96.7mm;" & vbCrLf
        tmpStr &= "width:199.5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleContainer table tr {" & vbCrLf
        tmpStr &= "vertical-align:top;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".detalleContainer table tr td{" & vbCrLf
        tmpStr &= "padding-bottom: 2mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerTotalHeader, .footerTotalDatos {" & vbCrLf
        tmpStr &= "text-align: right;" & vbCrLf
        tmpStr &= "left:133.721mm;" & vbCrLf
        tmpStr &= "top:215.4mm;" & vbCrLf
        tmpStr &= "font-size:3.2mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerTotalDatos {" & vbCrLf
        tmpStr &= "width: 25mm;" & vbCrLf
        tmpStr &= "left: 172mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerTotalHeader p, , .footerTotalDatos p {" & vbCrLf
        tmpStr &= "padding-bottom:4mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerTotalHeader .text, , .footerTotalDatos .text {" & vbCrLf
        tmpStr &= "font-size:3.6mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerBarcode{" & vbCrLf
        tmpStr &= "left:6mm;" & vbCrLf
        tmpStr &= "top:247.5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".barcodeContainer {" & vbCrLf
        tmpStr &= "height: 13mm;" & vbCrLf
        tmpStr &= "overflow:hidden;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerBarcode .barcode {" & vbCrLf
        tmpStr &= "font-family: ""Code 2 of 5 Interleaved"", sans-serif;" & vbCrLf
        tmpStr &= "text-align:left;" & vbCrLf
        tmpStr &= "font-size : 12mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerBarcode .barcodeNumber {" & vbCrLf
        tmpStr &= "font-size : 2.8mm;" & vbCrLf
        tmpStr &= "padding-left:12mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".footerBarcode small {" & vbCrLf
        tmpStr &= "font-weight: bold;" & vbCrLf
        tmpStr &= "font-style: italic;" & vbCrLf
        tmpStr &= "font-size: 2.1mm;" & vbCrLf
        tmpStr &= "padding-left: 2mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".comprobanteAutorizado{" & vbCrLf
        tmpStr &= "left:41mm;" & vbCrLf
        tmpStr &= "top:241mm;" & vbCrLf
        tmpStr &= "font-weight:bold;" & vbCrLf
        tmpStr &= "font-style:italic;" & vbCrLf
        tmpStr &= "font-size:3.19mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".paginacion {" & vbCrLf
        tmpStr &= "left:98mm;" & vbCrLf
        tmpStr &= "top:241mm;" & vbCrLf
        tmpStr &= "font-weight:bold;" & vbCrLf
        tmpStr &= "font-size:3.3mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".caeHeader, .caeDatos {" & vbCrLf
        tmpStr &= "font-weight: bold;" & vbCrLf
        tmpStr &= "left:128mm;" & vbCrLf
        tmpStr &= "top:239mm;" & vbCrLf
        tmpStr &= "text-align:right;" & vbCrLf
        tmpStr &= "font-size:3.3mm;F" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".caeHeader p, .caeDatos p {" & vbCrLf
        tmpStr &= "margin-bottom: 2.5mm;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".caeDatos {" & vbCrLf
        tmpStr &= "font-weight: normal;" & vbCrLf
        tmpStr &= "left:165mm;" & vbCrLf
        tmpStr &= "text-align:left;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "</style>" & vbCrLf
        Return tmpStr
    End Function



    Public Function templateFE(ByVal fe As AfipFactura, ByVal ex As AfipFacturaEX, ByVal gc As GlobalConfig) As String
        If fe.FechaEmision >= 20210203 Then
            RenderQR = True
        End If

        Dim tmpStr As String = ""
        tmpStr &= templateFEHead()

        Dim original As String = Me.FacturaBody(fe, ex, gc)
        Dim duplicado As String = original.Replace("#ORIGINAL#", "DUPLICADO").Replace("breakafter", "")
        original = original.Replace("#ORIGINAL#", "ORIGINAL")
        tmpStr &= original
        tmpStr &= duplicado

        tmpStr &= Me.templateFEFooter
        Return tmpStr
    End Function

    Public Function templateFEHead(Optional Filename As String = "") As String
        Dim tmpStr As String = ""
        tmpStr &= "<!doctype html>" & vbCrLf
        tmpStr &= "<html>" & vbCrLf
        tmpStr &= "<head>" & vbCrLf
        tmpStr &= "<meta charset=""utf-8""/>" & vbCrLf
        tmpStr &= "<title></title>" & vbCrLf
        tmpStr &= Me.CSSFull
        tmpStr &= "</head>" & vbCrLf
        tmpStr &= "<body>" & vbCrLf

        Return tmpStr
    End Function
    Public Function templateFEFooter() As String
        Dim tmpStr As String = ""
        tmpStr &= "</body>" & vbCrLf
        tmpStr &= "</html>" & vbCrLf
        Return tmpStr
    End Function

    Public Function FacturaBody(ByVal fe As AfipFactura, ByVal ex As AfipFacturaEX, ByVal gc As GlobalConfig) As String
        Dim numericoBarcode As String = ""
        Dim barcode As String = ""
        Dim tmpStr As String = ""

        If fe.ComprobanteTipo < AfipFactura.Tipo.RECIBO Then
            numericoBarcode = GenerarNumericBarcode(fe)
            barcode = StringToBarcode(numericoBarcode)
        End If

        tmpStr &= "<div class=""fe breakafter"">" & vbCrLf
        tmpStr &= "<div class=""blk header-blk"">" & vbCrLf
        tmpStr &= "<div class=""blk blk-inner header-orig-blk""></div>" & vbCrLf

        If fe.ComprobanteTipo < AfipFactura.Tipo.RECIBO Then
            tmpStr &= "<div class=""blk blk-inner header-type-blk""></div>" & vbCrLf
            tmpStr &= "<div class=""blk blk-inner header-logo-blk""></div>" & vbCrLf
        Else
            tmpStr &= "<div class=""blk blk-inner header-logo-blk"" style=""height:42mm;top:9mm""></div>" & vbCrLf
        End If



        tmpStr &= "</div>" & vbCrLf
        tmpStr &= "<div class=""blk date-blk""></div>" & vbCrLf
        tmpStr &= "<div class=""blk client-blk""></div>" & vbCrLf

        If fe.ComprobanteTipo < AfipFactura.Tipo.RECIBO Then
            tmpStr &= "<div class=""blk det-blk"">" & vbCrLf
        Else
            ' OCULTAR LA CABECERA DE DETALLES
            If fe.ComprobanteTipo = AfipFactura.Tipo.RECIBO Then
                tmpStr &= "<div class=""blk det-blk"" style=""display:none"">" & vbCrLf
            Else
                tmpStr &= "<div class=""blk det-blk"">" & vbCrLf
            End If

        End If

        tmpStr &= "<div class=""blk blk-inner det-codigo-blk""></div>" & vbCrLf
        tmpStr &= "<div class=""blk blk-inner det-producto-blk""></div>" & vbCrLf
        tmpStr &= "<div class=""blk blk-inner det-cantidad-blk""></div>" & vbCrLf
        tmpStr &= "<div class=""blk blk-inner det-medida-blk""></div>" & vbCrLf
        tmpStr &= "<div class=""blk blk-inner det-precio-blk""></div>" & vbCrLf
        tmpStr &= "<div class=""blk blk-inner det-bonif-blk""></div>" & vbCrLf
        tmpStr &= "<div class=""blk blk-inner det-imp-blk""></div>" & vbCrLf
        tmpStr &= "</div>" & vbCrLf

        If fe.ComprobanteTipo < AfipFactura.Tipo.RECIBO Then
            tmpStr &= "<div class=""blk price-blk""></div>" & vbCrLf
        Else
            tmpStr &= "<div class=""blk price-blk"" style=""height:38mm""></div>" & vbCrLf
        End If

        tmpStr &= $"<div class=""content originalDuplicado"">#ORIGINAL#</div>" & vbCrLf

        If fe.ComprobanteTipo < AfipFactura.Tipo.RECIBO Then
            tmpStr &= "<div class=""content tipoComprobante"">" & vbCrLf
        End If

        Select Case fe.ComprobanteTipo
            Case AfipFactura.Tipo.FACTURA_A
                tmpStr &= $"<p class=""tipo"">A</p>" & vbCrLf
                tmpStr &= $"<p class=""tipoCodigo"">COD. 01</p>" & vbCrLf
            Case AfipFactura.Tipo.FACTURA_B
                tmpStr &= $"<p class=""tipo"">B</p>" & vbCrLf
                tmpStr &= $"<p class=""tipoCodigo"">COD. 06</p>" & vbCrLf
            Case AfipFactura.Tipo.FACTURA_C
                tmpStr &= $"<p class=""tipo"">C</p>" & vbCrLf
                tmpStr &= $"<p class=""tipoCodigo"">COD. 11</p>" & vbCrLf
            Case AfipFactura.Tipo.NOTA_CREDITO_C
                tmpStr &= $"<p class=""tipo"">C</p>" & vbCrLf
                tmpStr &= $"<p class=""tipoCodigo"">COD. 13</p>" & vbCrLf
            Case AfipFactura.Tipo.NOTA_DEBITO_C
                tmpStr &= $"<p class=""tipo"">C</p>" & vbCrLf
                tmpStr &= $"<p class=""tipoCodigo"">COD. 12</p>" & vbCrLf
            Case AfipFactura.Tipo.RECIBO
                ' NO HACER NADA
                'tmpStr &= $"<p class=""tipo"">X</p>" & vbCrLf
                'tmpStr &= $"<p class=""tipoCodigo"">COD. 00</p>" & vbCrLf

        End Select

        If fe.ComprobanteTipo < AfipFactura.Tipo.RECIBO Then
            tmpStr &= "</div>" & vbCrLf
        End If

        Select Case gc.FeTipoLogo
            Case 0
                tmpStr &= $"<div class=""content logo"">{gc.NombreFantasia}</div>" & vbCrLf
            Case 1
                tmpStr &= $"<div class=""content logo""><img style=""width:65mm;height:auto"" src=""data:image/jpeg;base64,{ImageToBase64(gc.FeLogo)}""/></div>" & vbCrLf
        End Select


        If fe.ComprobanteTipo = AfipFactura.Tipo.RECIBO Then
            If fe.Detalles(0).Codigo = "666333" Then
                fe.Detalles.RemoveAt(0)
                tmpStr &= "<div class=""content headerComprobante"" style=""top: 41.252mm;"">" & vbCrLf
                tmpStr &= "<p class=""text"" style=""font-size:3.3mm""><b>SECCION REGISTRO DE CONTRATOS y CERTIFICADOS</b></p>" & vbCrLf
            Else
                tmpStr &= "<div class=""content headerComprobante"">" & vbCrLf
            End If

        Else
            tmpStr &= "<div class=""content headerComprobante"">" & vbCrLf
        End If

        tmpStr &= $"<p class=""text""><b>Razón Social:</b> {ex.RazonSocialEmisor}</p>" & vbCrLf
        tmpStr &= $"<p class=""text""><b>Domicilio Comercial:</b> {ex.DomicilioEmisor}</p>" & vbCrLf
        tmpStr &= $"<p class=""text""><b>Condición frente al IVA: &nbsp;&nbsp;IVA EXENTO</b></p>" & vbCrLf
        tmpStr &= "</div>" & vbCrLf
        tmpStr &= "<div class=""content detalleComprobante"">" & vbCrLf
        If fe.ComprobanteTipo = AfipFactura.Tipo.FACTURA_A Or
            fe.ComprobanteTipo = AfipFactura.Tipo.FACTURA_B Or
            fe.ComprobanteTipo = AfipFactura.Tipo.FACTURA_C Then

            tmpStr &= "<p class=""tipo""><b>FACTURA</b></p>" & vbCrLf
        End If
        If fe.ComprobanteTipo = AfipFactura.Tipo.NOTA_DEBITO_A Or
            fe.ComprobanteTipo = AfipFactura.Tipo.NOTA_DEBITO_B Or
            fe.ComprobanteTipo = AfipFactura.Tipo.NOTA_DEBITO_C Then

            tmpStr &= "<p class=""tipo""><b>NOTA DE DEBITO</b></p>" & vbCrLf
        End If
        If fe.ComprobanteTipo = AfipFactura.Tipo.NOTA_CREDITO_A Or
            fe.ComprobanteTipo = AfipFactura.Tipo.NOTA_CREDITO_B Or
            fe.ComprobanteTipo = AfipFactura.Tipo.NOTA_CREDITO_C Then

            tmpStr &= "<p class=""tipo""><b>NOTA DE CREDITO</b></p>" & vbCrLf
        End If
        If fe.ComprobanteTipo = AfipFactura.Tipo.RECIBO_A Or
            fe.ComprobanteTipo = AfipFactura.Tipo.RECIBO_B Or
            fe.ComprobanteTipo = AfipFactura.Tipo.RECIBO_C Then

            tmpStr &= "<p class=""tipo""><b>RECIBO</b></p>" & vbCrLf
        End If

        If fe.ComprobanteTipo >= AfipFactura.Tipo.RECIBO Then
            Dim strTipoComprobante As String
            Select Case fe.ComprobanteTipo
                Case AfipFactura.Tipo.RECIBO
                    strTipoComprobante = "RECIBO"
                Case AfipFactura.Tipo.NOTA_CREDITO_X
                    strTipoComprobante = "NOTA DE CRÉDITO"
                Case AfipFactura.Tipo.NOTA_DEBITO_X
                    strTipoComprobante = "NOTA DE DÉBITO"
            End Select
            tmpStr &= $"<p class=""tipo""  style=""font-size:18px""><b>{strTipoComprobante}<span style=""font-size:9px;margin-bottom:0""><br>DOCUMENTO NO VALIDO COMO FACTURA</span></b></p>" & vbCrLf
            tmpStr &= $"<p class=""text""style=""margin-bottom:3.3mm;margin-top:-9px""><b>Punto de Venta: &nbsp;{fe.PuntoVenta.ToString.PadLeft(4, "0")}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Comp. Nro: &nbsp;{fe.Numero.ToString.PadLeft(8, "0")}</b></p>" & vbCrLf
        Else
            tmpStr &= $"<p class=""text""style=""margin-bottom:3.3mm""><b>Punto de Venta: &nbsp;{fe.PuntoVenta.ToString.PadLeft(4, "0")}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Comp. Nro: &nbsp;{fe.Numero.ToString.PadLeft(8, "0")}</b></p>" & vbCrLf
        End If


        tmpStr &= $"<p class=""text""style=""margin-bottom:5mm""><b>Fecha de Emisión: &nbsp;{fromISO8601(fe.FechaEmision)}</b></p>" & vbCrLf
        tmpStr &= $"<p class=""text""><b>CUIT:</b> {ToCuitConGuiones(fe.CuitEmisor)}</p>" & vbCrLf
        tmpStr &= $"<p class=""text""><b>Ingresos Brutos:</b> {ToCuitConGuiones(gc.IngresosBrutos)}</p>" & vbCrLf
        tmpStr &= $"<p class=""text""><b>Fecha de Inicio de Actividades:</b> {fromISO8601(gc.FechaInicio)}</p>" & vbCrLf
        tmpStr &= "</div>" & vbCrLf
        If fe.Concept = AfipFactura.Concepto.SERVICIOS Or fe.Concept = AfipFactura.Concepto.PRODUCTOS_Y_SERVICIOS Then
            tmpStr &= "<div class=""content fechasComprobante"">" & vbCrLf
            tmpStr &= $"<b>Período Facturado Desde:</b> {fromISO8601(fe.FechaServicioDesde)} <b style=""margin-left:20mm"">Hasta:</b> {fromISO8601(fe.FechaServicioHasta)} <b style=""margin-left:17mm"">Fecha de Vto. para el pago: </b> {fromISO8601(fe.FechaVencimientoPago)}" & vbCrLf
            tmpStr &= "</div>" & vbCrLf
        End If
        tmpStr &= "<div class=""content detalleReceptor"">" & vbCrLf
        Select Case fe.DocumentoTipo
            Case AfipFactura.Documento.CUIT
                tmpStr &= $"<p class=""text""><b>CUIT:</b> {ToCuitConGuiones(fe.DocumentoNumero)}</p>" & vbCrLf
            Case AfipFactura.Documento.DNI
                tmpStr &= $"<p class=""text""><b>DNI:</b> {ToCuitConGuiones(fe.DocumentoNumero)}</p>" & vbCrLf
            Case AfipFactura.Documento.CUIL
                tmpStr &= $"<p class=""text""><b>CUIL:</b> {ToCuitConGuiones(fe.DocumentoNumero)}</p>" & vbCrLf
            Case Else
                tmpStr &= $"<p class=""text""><b>CUIT:</b> {ToCuitConGuiones(fe.DocumentoNumero)}</p>" & vbCrLf
        End Select
        tmpStr &= $"<p class=""text""><b>Condición frente al IVA:</b> {ex.CondicionFiscalStringReceptor}</p>" & vbCrLf

        Dim condicionPago As String = ""
        If ex.CondicionContado Then
            condicionPago = "Contado"
        End If

        If ex.CondicionTarjetaDebito Then
            If condicionPago.Length > 0 Then condicionPago &= " - "
            condicionPago &= "T. Débito"
        End If

        If ex.CondicionTarjetaCredito Then
            If condicionPago.Length > 0 Then condicionPago &= " - "
            condicionPago &= "T. Crédito"
        End If

        If ex.CondicionCuentaCorriente Then
            If condicionPago.Length > 0 Then condicionPago &= " - "
            condicionPago &= "C. Corriente"
        End If

        If ex.CondicionCheque Then
            If condicionPago.Length > 0 Then condicionPago &= " - "
            condicionPago &= "Cheque"
        End If

        If ex.CondicionTicket Then
            If condicionPago.Length > 0 Then condicionPago &= " - "
            condicionPago &= "Ticket"
        End If

        If ex.CondicionOtra Then
            If condicionPago.Length > 0 Then condicionPago &= " - "
            condicionPago &= "Otra"
        End If

        tmpStr &= $"<p class=""text""><b>Condición de venta:</b> {condicionPago}</p>" & vbCrLf
        tmpStr &= "</div>" & vbCrLf
        tmpStr &= "<div class=""content detalleReceptorB"">" & vbCrLf
        tmpStr &= $"<p class=""text""><b>Apellido y Nombre / Razón Social:</b> {ex.RazonSocialReceptor}</p>" & vbCrLf
        tmpStr &= $"<p class=""text""><b>Domicilio:</b> {ex.DomicilioReceptor}</p>" & vbCrLf
        tmpStr &= "</div>" & vbCrLf

        If fe.ComprobanteTipo < AfipFactura.Tipo.RECIBO Or fe.ComprobanteTipo = AfipFactura.Tipo.NOTA_CREDITO_X Or fe.ComprobanteTipo = AfipFactura.Tipo.NOTA_DEBITO_X Then
            tmpStr &= "<div class=""content detalleHeaders"">" & vbCrLf
        Else
            tmpStr &= "<div class=""content detalleHeaders"" style=""display:none"">" & vbCrLf
        End If

        tmpStr &= "<div><b>Código</b></div>" & vbCrLf
        tmpStr &= "<div style=""left:25mm""><b>Producto / Servicio</b></div>" & vbCrLf
        tmpStr &= "<div style=""left:67mm""><b>Cantidad</b></div>" & vbCrLf
        tmpStr &= "<div style=""left:85mm""><b>U. Medida</b></div>" & vbCrLf
        tmpStr &= "<div style=""left:105mm""><b>Precio Unit.</b></div>" & vbCrLf
        tmpStr &= "<div style=""left:128.5mm""><b style=""font-size:2.25mm"">% Bonif.</b></div>" & vbCrLf
        tmpStr &= "<div style=""left:144mm""><b>Imp. Bonif.</b></div>" & vbCrLf
        tmpStr &= "<div style=""left:175mm""><b>Subtotal</b></div>" & vbCrLf
        tmpStr &= "</div>" & vbCrLf
        tmpStr &= "<div class=""content detalleContainer"">" & vbCrLf
        tmpStr &= "<table>" & vbCrLf

        Dim importeTotal As Decimal = 0

        If fe.ComprobanteTipo < AfipFactura.Tipo.RECIBO Then
            For Each det As AfipFacturaDetalle In fe.Detalles
                tmpStr &= "<tr>" & vbCrLf
                tmpStr &= $"<td style=""width:14.204mm;padding-left:1mm"">{det.Codigo}</td>" & vbCrLf
                tmpStr &= $"<td style=""width:49.808mm;padding-left:1mm"">{det.ProductoServicio}</td>" & vbCrLf
                tmpStr &= $"<td style=""width:22.754mm;text-align:right"">{det.Cantidad}</td>" & vbCrLf
                tmpStr &= $"<td style=""width:15.015mm;text-align:center"">{det.UnidadMedida.ToString.Replace("_", " ").ToLower}</td>" & vbCrLf
                tmpStr &= $"<td style=""width:28.377mm;text-align:right"">{ToAfipMoneyFormat(det.PrecioUnitario)}</td>" & vbCrLf
                tmpStr &= $"<td style=""width:11.377mm;text-align:center"">{ToAfipMoneyFormat(det.BonificacionPercent)}</td>" & vbCrLf
                tmpStr &= $"<td style=""width:25.4mm;text-align:right"">{ToAfipMoneyFormat(((det.PrecioUnitario * det.Cantidad) * det.BonificacionPercent) / 100)}</td>" & vbCrLf
                Dim subtotal As Decimal = (det.PrecioUnitario * det.Cantidad) - (((det.PrecioUnitario * det.Cantidad) * det.BonificacionPercent) / 100)
                tmpStr &= $"<td style=""width:32mm;text-align:right;padding-right:1mm"">{ToAfipMoneyFormat(subtotal)}</td>" & vbCrLf
                tmpStr &= $"</tr>" & vbCrLf
                importeTotal += subtotal
            Next
        Else
            If fe.ComprobanteTipo = AfipFactura.Tipo.RECIBO Then
                tmpStr &= $"<tr><td style=""width:129.204mm;padding-left:1mm""><b>Recibi(mos) la suma de:</b> {Utils.ToMoneyFormat(fe.ImporteTotal)}</td></tr>" & vbCrLf
                tmpStr &= "<tr><td style=""width:129.204mm;padding-left:1mm""><b>en concepto de:</b></td></tr>" & vbCrLf
                For Each det As AfipFacturaDetalle In fe.Detalles
                    tmpStr &= $"<tr><td style=""width:129.204mm;padding-left:1mm"">{det.ProductoServicio}</td></tr>" & vbCrLf
                    Dim subtotal As Decimal = (det.PrecioUnitario * det.Cantidad) - (((det.PrecioUnitario * det.Cantidad) * det.BonificacionPercent) / 100)
                    importeTotal += subtotal
                Next
            Else
                For Each det As AfipFacturaDetalle In fe.Detalles
                    tmpStr &= "<tr>" & vbCrLf
                    tmpStr &= $"<td style=""width:14.204mm;padding-left:1mm"">{det.Codigo}</td>" & vbCrLf
                    tmpStr &= $"<td style=""width:49.808mm;padding-left:1mm"">{det.ProductoServicio}</td>" & vbCrLf
                    tmpStr &= $"<td style=""width:22.754mm;text-align:right"">{det.Cantidad}</td>" & vbCrLf
                    tmpStr &= $"<td style=""width:15.015mm;text-align:center"">{det.UnidadMedida.ToString.Replace("_", " ").ToLower}</td>" & vbCrLf
                    tmpStr &= $"<td style=""width:28.377mm;text-align:right"">{ToAfipMoneyFormat(det.PrecioUnitario)}</td>" & vbCrLf
                    tmpStr &= $"<td style=""width:11.377mm;text-align:center"">{ToAfipMoneyFormat(det.BonificacionPercent)}</td>" & vbCrLf
                    tmpStr &= $"<td style=""width:25.4mm;text-align:right"">{ToAfipMoneyFormat(((det.PrecioUnitario * det.Cantidad) * det.BonificacionPercent) / 100)}</td>" & vbCrLf
                    Dim subtotal As Decimal = (det.PrecioUnitario * det.Cantidad) - (((det.PrecioUnitario * det.Cantidad) * det.BonificacionPercent) / 100)
                    tmpStr &= $"<td style=""width:32mm;text-align:right;padding-right:1mm"">{ToAfipMoneyFormat(subtotal)}</td>" & vbCrLf
                    tmpStr &= $"</tr>" & vbCrLf
                    importeTotal += subtotal
                Next
            End If
        End If




        tmpStr &= "</table>" & vbCrLf
        tmpStr &= "</div>" & vbCrLf

        If fe.ComprobanteTipo < AfipFactura.Tipo.RECIBO Or fe.ComprobanteTipo = AfipFactura.Tipo.NOTA_CREDITO_X Then
            tmpStr &= "<div class=""content footerTotalHeader"">" & vbCrLf
            tmpStr &= "<p><b>Subtotal: $</b></p>" & vbCrLf
            tmpStr &= "<p><b>Importe Otros Tributos: $</b></p>" & vbCrLf
            tmpStr &= "<p class=""text""><b>Importe Total: $</b></p>" & vbCrLf
            tmpStr &= "</div>" & vbCrLf
            tmpStr &= "<div class=""content footerTotalDatos"">" & vbCrLf
            tmpStr &= $"<p><b>{ToAfipMoneyFormat(importeTotal)}</b></p>" & vbCrLf
            tmpStr &= $"<p><b>{ToAfipMoneyFormat(fe.ImporteTributo)}</b></p>" & vbCrLf
            tmpStr &= $"<p Class=""text""><b>{ToAfipMoneyFormat(importeTotal)}</b></p>" & vbCrLf
            tmpStr &= "</div>" & vbCrLf
        Else
            tmpStr &= "<div class=""content footerTotalHeader"" style=""top:205mm"">" & vbCrLf
            tmpStr &= "<p><b>Subtotal: $</b></p>" & vbCrLf
            tmpStr &= "<p><b>Bonif: </b>% 0 &nbsp;&nbsp;<b>Importe Bonif: $</b></p>" & vbCrLf
            tmpStr &= "<p><b>Subtotal c/Bonif.: $</b></p>" & vbCrLf
            tmpStr &= "<p><b>Importe Otros Tributos: $</b></p>" & vbCrLf
            tmpStr &= "<p class=""text""><b>Importe Total: $</b></p>" & vbCrLf
            tmpStr &= "</div>" & vbCrLf
            tmpStr &= "<div class=""content footerTotalDatos"" style=""top:205mm"">" & vbCrLf
            tmpStr &= $"<p><b>{ToAfipMoneyFormat(If(importeTotal > 0, importeTotal, fe.ImporteTotal))}</b></p>" & vbCrLf
            tmpStr &= "<p><b>0,00</b></p>" & vbCrLf
            tmpStr &= $"<p><b>{ToAfipMoneyFormat(If(importeTotal > 0, importeTotal, fe.ImporteTotal))}</b></p>" & vbCrLf
            tmpStr &= "<p><b>0,00</b></p>" & vbCrLf
            tmpStr &= $"<p Class=""text""><b>{ToAfipMoneyFormat(If(importeTotal > 0, importeTotal, fe.ImporteTotal))}</b></p>" & vbCrLf
            tmpStr &= "</div>" & vbCrLf

        End If


        If fe.ComprobanteTipo >= AfipFactura.Tipo.RECIBO Then
            tmpStr &= "</div>" & vbCrLf
            tmpStr &= "</div>" & vbCrLf
        Else
            tmpStr &= "<div class=""content footerBarcode"">" & vbCrLf
            tmpStr &= "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAZ0AAABxCAQAAACkJ7ZkAAAACXBIWXMAAAsTAAALEwEAmpwYAAADGWlDQ1BQaG90b3Nob3AgSUNDIHByb2ZpbGUAAHjaY2BgnuDo4uTKJMDAUFBUUuQe5BgZERmlwH6egY2BmYGBgYGBITG5uMAxIMCHgYGBIS8/L5UBA3y7xsDIwMDAcFnX0cXJlYE0wJpcUFTCwMBwgIGBwSgltTiZgYHhCwMDQ3p5SUEJAwNjDAMDg0hSdkEJAwNjAQMDg0h2SJAzAwNjCwMDE09JakUJAwMDg3N+QWVRZnpGiYKhpaWlgmNKflKqQnBlcUlqbrGCZ15yflFBflFiSWoKAwMD1A4GBgYGXpf8EgX3xMw8BUNTVQYqg4jIKAX08EGIIUByaVEZhMXIwMDAIMCgxeDHUMmwiuEBozRjFOM8xqdMhkwNTJeYNZgbme+y2LDMY2VmzWa9yubEtoldhX0mhwBHJycrZzMXM1cbNzf3RB4pnqW8xryH+IL5nvFXCwgJrBZ0E3wk1CisKHxYJF2UV3SrWJw4p/hWiRRJYcmjUhXSutJPZObIhsoJyp2V71HwUeRVvKA0RTlKRUnltepWtUZ1Pw1Zjbea+7QmaqfqWOsK6b7SO6I/36DGMMrI0ljS+LfJPdPDZivM+y0qLBOtfKwtbFRtRexY7L7aP3e47XjB6ZjzXpetruvdVrov9VjkudBrgfdCn8W+y/xW+a8P2Bq4N+hY8PmQW6HPwr5EMEUKRilFG8e4xUbF5cW3JMxO3Jx0Nvl5KlOaXLpNRlRmVdas7D059/KY8tULfAqLi2YXHy55WyZR7lJRWDmv6mz131q9uvj6SQ3HGn83G7Skt85ru94h2Ond1d59uJehz76/bsK+if8nO05pnXpiOu+M4JmzZj2aozW3ZN6+BVwLwxYtXvxxqcOyCcsfrjRe1br65lrddU3rb2402NSx+cFWq21Tt3/Y6btr1R6Oven7jh9QP9h56PURv6Obj4ufqD355LT3mS3nZM+3X/h0Ke7yqasW15bdEL3ZeuvrnfS7N+/7PDjwyPTx6qeKz2a+EHzZ9Zr5Td3bn+9LP3z6VPD53de8b+9+5P/88Lv4z7d/Vf//AwAqvx2K829RWwAAOGlpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+Cjx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuNi1jMDY3IDc5LjE1Nzc0NywgMjAxNS8wMy8zMC0yMzo0MDo0MiAgICAgICAgIj4KICAgPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4KICAgICAgPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIKICAgICAgICAgICAgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIgogICAgICAgICAgICB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8iCiAgICAgICAgICAgIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIKICAgICAgICAgICAgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iCiAgICAgICAgICAgIHhtbG5zOnN0RXZ0PSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VFdmVudCMiCiAgICAgICAgICAgIHhtbG5zOnRpZmY9Imh0dHA6Ly9ucy5hZG9iZS5jb20vdGlmZi8xLjAvIgogICAgICAgICAgICB4bWxuczpleGlmPSJodHRwOi8vbnMuYWRvYmUuY29tL2V4aWYvMS4wLyI+CiAgICAgICAgIDx4bXA6Q3JlYXRvclRvb2w+QWRvYmUgUGhvdG9zaG9wIENDIDIwMTUgKFdpbmRvd3MpPC94bXA6Q3JlYXRvclRvb2w+CiAgICAgICAgIDx4bXA6Q3JlYXRlRGF0ZT4yMDE5LTA5LTMwVDE5OjUxOjM5LTAzOjAwPC94bXA6Q3JlYXRlRGF0ZT4KICAgICAgICAgPHhtcDpNb2RpZnlEYXRlPjIwMTktMDktMzBUMTk6NTU6MzctMDM6MDA8L3htcDpNb2RpZnlEYXRlPgogICAgICAgICA8eG1wOk1ldGFkYXRhRGF0ZT4yMDE5LTA5LTMwVDE5OjU1OjM3LTAzOjAwPC94bXA6TWV0YWRhdGFEYXRlPgogICAgICAgICA8ZGM6Zm9ybWF0PmltYWdlL3BuZzwvZGM6Zm9ybWF0PgogICAgICAgICA8cGhvdG9zaG9wOkNvbG9yTW9kZT4xPC9waG90b3Nob3A6Q29sb3JNb2RlPgogICAgICAgICA8cGhvdG9zaG9wOklDQ1Byb2ZpbGU+RG90IEdhaW4gMTUlPC9waG90b3Nob3A6SUNDUHJvZmlsZT4KICAgICAgICAgPHhtcE1NOkluc3RhbmNlSUQ+eG1wLmlpZDphMDJjZmU0ZC0yNmQ1LTVhNGMtOGIyYi02NmE4YTliN2Q2Y2U8L3htcE1NOkluc3RhbmNlSUQ+CiAgICAgICAgIDx4bXBNTTpEb2N1bWVudElEPnhtcC5kaWQ6YTAyY2ZlNGQtMjZkNS01YTRjLThiMmItNjZhOGE5YjdkNmNlPC94bXBNTTpEb2N1bWVudElEPgogICAgICAgICA8eG1wTU06T3JpZ2luYWxEb2N1bWVudElEPnhtcC5kaWQ6YTAyY2ZlNGQtMjZkNS01YTRjLThiMmItNjZhOGE5YjdkNmNlPC94bXBNTTpPcmlnaW5hbERvY3VtZW50SUQ+CiAgICAgICAgIDx4bXBNTTpIaXN0b3J5PgogICAgICAgICAgICA8cmRmOlNlcT4KICAgICAgICAgICAgICAgPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDphY3Rpb24+Y3JlYXRlZDwvc3RFdnQ6YWN0aW9uPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6aW5zdGFuY2VJRD54bXAuaWlkOmEwMmNmZTRkLTI2ZDUtNWE0Yy04YjJiLTY2YThhOWI3ZDZjZTwvc3RFdnQ6aW5zdGFuY2VJRD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OndoZW4+MjAxOS0wOS0zMFQxOTo1MTozOS0wMzowMDwvc3RFdnQ6d2hlbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OnNvZnR3YXJlQWdlbnQ+QWRvYmUgUGhvdG9zaG9wIENDIDIwMTUgKFdpbmRvd3MpPC9zdEV2dDpzb2Z0d2FyZUFnZW50PgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgPC9yZGY6U2VxPgogICAgICAgICA8L3htcE1NOkhpc3Rvcnk+CiAgICAgICAgIDx0aWZmOk9yaWVudGF0aW9uPjE8L3RpZmY6T3JpZW50YXRpb24+CiAgICAgICAgIDx0aWZmOlhSZXNvbHV0aW9uPjcyMDAwMC8xMDAwMDwvdGlmZjpYUmVzb2x1dGlvbj4KICAgICAgICAgPHRpZmY6WVJlc29sdXRpb24+NzIwMDAwLzEwMDAwPC90aWZmOllSZXNvbHV0aW9uPgogICAgICAgICA8dGlmZjpSZXNvbHV0aW9uVW5pdD4yPC90aWZmOlJlc29sdXRpb25Vbml0PgogICAgICAgICA8ZXhpZjpDb2xvclNwYWNlPjY1NTM1PC9leGlmOkNvbG9yU3BhY2U+CiAgICAgICAgIDxleGlmOlBpeGVsWERpbWVuc2lvbj40MTM8L2V4aWY6UGl4ZWxYRGltZW5zaW9uPgogICAgICAgICA8ZXhpZjpQaXhlbFlEaW1lbnNpb24+MTEzPC9leGlmOlBpeGVsWURpbWVuc2lvbj4KICAgICAgPC9yZGY6RGVzY3JpcHRpb24+CiAgIDwvcmRmOlJERj4KPC94OnhtcG1ldGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgCjw/eHBhY2tldCBl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"" class=""logo-afip"" />" & vbCrLf


            If fe.FechaEmision >= 20210203 Then
                RenderQR = True
                tmpStr &= $"<img class=""qrcode"" src=""data:image/png;base64,{GenerarQR(fe)}""/>" & vbCrLf
            End If

            tmpStr &= "<small>Esta Administración Federal no se responsabiliza por los datos ingresados en el detalle de la operación</small>" & vbCrLf
            tmpStr &= "<div class=""barcodeContainer"">" & vbCrLf
            tmpStr &= $"<p class=""barcode"">{barcode}</p>" & vbCrLf
            tmpStr &= $"<p class=""barcode"" style=""margin:0;line-height:0.65"">{barcode}</p>" & vbCrLf
            tmpStr &= "</div>" & vbCrLf
            tmpStr &= $"<p class=""barcodeNumber"">{numericoBarcode}</p>" & vbCrLf
            tmpStr &= "</div>" & vbCrLf
            tmpStr &= "<div class=""content comprobanteAutorizado"">Comprobante Autorizado</div>" & vbCrLf
            tmpStr &= "<div class=""content paginacion"">Pág. 1/1</div>" & vbCrLf
            tmpStr &= "<div class=""content caeHeader"">" & vbCrLf
            tmpStr &= "<p>CAE Nº:</p>" & vbCrLf
            tmpStr &= "<p>Fecha de Vto. de CAE:</p>" & vbCrLf
            tmpStr &= "</div>" & vbCrLf
            tmpStr &= "<div class=""content caeDatos"">" & vbCrLf
            tmpStr &= $"<p>{fe.Cae}</p>" & vbCrLf
            tmpStr &= $"<p>{fromISO8601(fe.FechaVencimiento)}</p>" & vbCrLf
            tmpStr &= "</div>" & vbCrLf
            tmpStr &= "</div>" & vbCrLf
        End If


        Return tmpStr
    End Function

    Public Shared Function CalcularDigitoVerificadorBarcode(ByVal tabla As String) As Integer
        Dim pares As Integer = 0
        Dim impares As Integer = 0
        Dim esPar As Boolean = False

        For Each caracter As Char In tabla
            If esPar Then
                pares += CInt(caracter.ToString)
            Else
                impares += CInt(caracter.ToString)
            End If
            esPar = Not esPar
        Next

        pares *= 3
        pares += impares

        Dim tmp As Integer = (pares / 10)
        tmp = tmp * 10


        Return Math.Abs(tmp - pares)
    End Function

    Public Shared Function GenerarNumericBarcode(ByVal fe As AfipFactura) As String
        Dim tablaCalculo As String
        If fe.ComprobanteTipo < AfipFactura.Tipo.NOTA_VENTA_CONTADO_B Then
            tablaCalculo = fe.CuitEmisor.ToString & "0" & fe.ComprobanteTipo & fe.PuntoVenta.ToString.PadLeft(4, "0") & fe.Cae.ToString & fe.FechaVencimiento.ToString
        Else
            tablaCalculo = fe.CuitEmisor.ToString & fe.ComprobanteTipo & fe.PuntoVenta.ToString.PadLeft(4, "0") & fe.Cae.ToString & fe.FechaVencimiento.ToString
        End If
        tablaCalculo = tablaCalculo & CalcularDigitoVerificadorBarcode(tablaCalculo)

        Return tablaCalculo
    End Function

    Public Function StringToBarcode(ByVal numericBarcode As String) As String
        If numericBarcode.Length <> 40 Then
            Return ""
        End If

        Dim output As String = "É"
        For i As Integer = 1 To numericBarcode.Length Step 2
            If (CInt(numericBarcode.Substring(i - 1, 2))) > 93 Then
                output &= Chr(CInt(numericBarcode.Substring(i - 1, 2)) + 101)
            Else
                output &= Chr(CInt(numericBarcode.Substring(i - 1, 2)) + 33)
            End If
        Next
        output &= "Ê"

        Return output
    End Function


    Public Function RenderMultiple(ByVal lstFe As List(Of AfipFactura), ByVal lstEx As List(Of AfipFacturaEX), ByVal gc As GlobalConfig, Optional forceBreak As Boolean = False) As String
        Dim tmpStr As String = ""
        If lstFe.Count <> lstEx.Count Then
            Return ""
        End If


        If Utils.DateTo8601(Now.Date) >= 20210203 Then
            RenderQR = True
        End If

        tmpStr &= Me.templateFEHead

        Dim currentIndex As Integer = 0
        For Each fe As AfipFactura In lstFe


            Dim original As String = Me.FacturaBody(fe, lstEx(currentIndex), gc)
            ' BUMP
            Dim duplicado As String = ""
            If forceBreak Then
                If currentIndex = lstFe.Count - 1 Then
                    duplicado = original.Replace("#ORIGINAL#", "DUPLICADO").Replace("breakafter", "")
                Else
                    duplicado = original.Replace("#ORIGINAL#", "DUPLICADO")
                End If
            Else
                If currentIndex = lstFe.Count - 1 Then
                    duplicado = original.Replace("#ORIGINAL#", "DUPLICADO")
                Else
                    duplicado = original.Replace("#ORIGINAL#", "DUPLICADO").Replace("breakafter", "")
                End If

            End If
            original = original.Replace("#ORIGINAL#", "ORIGINAL")
            tmpStr &= original
            tmpStr &= duplicado
            currentIndex += 1
        Next

        tmpStr &= Me.templateFEFooter

        Return tmpStr
    End Function


    Public Function ToPDF(ByVal fe As AfipFactura, ByVal fex As AfipFacturaEX, ByVal gc As GlobalConfig, ByVal path As String) As Boolean
        Dim tmpHTML As String = ""
        Try
            tmpHTML = Me.templateFE(fe, fex, gc)
            Dim pdfRender As New HtmlToPdf
            Dim doc As PdfDocument = pdfRender.ConvertHtmlString(tmpHTML)
            doc.Save(path)
            doc.Close()
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function


    Public Function GenerarQR(ByVal fe As AfipFactura) As String
        Dim afipQR As New AfipQREncoder

        With afipQR
            .CodAut = fe.Cae
            .Ctz = fe.MonedaCotizacion
            .Cuit = fe.CuitEmisor
            .Fecha = Utils.Int8601ToDate(fe.FechaEmision)
            .Importe = fe.ImporteTotal
            .Moneda = fe.Moneda
            .NroCmp = fe.Numero
            .PtoVta = fe.PuntoVenta
            .TipoCmp = fe.ComprobanteTipo
            .TipoCodAut = AfipQREncoder.TIPO_CAE
        End With

        Dim qrData As New QRCodeData(1)
        Dim qrgen As New QRCodeGenerator()


        qrData = qrgen.CreateQrCode(afipQR.GenerarURLEncoded, QRCodeGenerator.ECCLevel.L)
        Dim qr As New Base64QRCode(qrData)


        Return qr.GetGraphic(10)
    End Function


End Class
