Module TEMPLATES
    Public Function PlantillaFE(ByVal fe As AfipFactura, ByVal fex As AfipFacturaEX, ByVal gc As GlobalConfig) As String
        Dim tmpStr As String = ""
        tmpStr &= "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrLf
        tmpStr &= "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf
        tmpStr &= "<head>" & vbCrLf
        tmpStr &= "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"" />" & vbCrLf
        tmpStr &= "<meta name=""viewport"" content=""width=device-width, initial-scale=1"" />" & vbCrLf
        tmpStr &= "<title>Oxygen Confirm</title>" & vbCrLf
        tmpStr &= "<style type=""text/css"">" & vbCrLf
        tmpStr &= "/* Take care of image borders and formatting, client hacks */" & vbCrLf
        tmpStr &= "img { max-width: 600px; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic;}" & vbCrLf
        tmpStr &= "a img { border: none; }" & vbCrLf
        tmpStr &= "table { border-collapse: collapse !important;}" & vbCrLf
        tmpStr &= "#outlook a { padding:0; }" & vbCrLf
        tmpStr &= ".ReadMsgBody { width: 100%; }" & vbCrLf
        tmpStr &= ".ExternalClass { width: 100%; }" & vbCrLf
        tmpStr &= ".backgroundTable { margin: 0 auto; padding: 0; width: 100% !important; }" & vbCrLf
        tmpStr &= "table td { border-collapse: collapse; }" & vbCrLf
        tmpStr &= ".ExternalClass * { line-height: 115%; }" & vbCrLf
        tmpStr &= ".container-for-gmail-android { min-width: 600px; }" & vbCrLf
        tmpStr &= "/* General styling */" & vbCrLf
        tmpStr &= "* {" & vbCrLf
        tmpStr &= "font-family: Helvetica, Arial, sans-serif;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "body {" & vbCrLf
        tmpStr &= "-webkit-font-smoothing: antialiased;" & vbCrLf
        tmpStr &= "-webkit-text-size-adjust: none;" & vbCrLf
        tmpStr &= "width: 100% !important;" & vbCrLf
        tmpStr &= "margin: 0 !important;" & vbCrLf
        tmpStr &= "height: 100%;" & vbCrLf
        tmpStr &= "color: #676767;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td {" & vbCrLf
        tmpStr &= "font-family: Helvetica, Arial, sans-serif;" & vbCrLf
        tmpStr &= "font-size: 14px;" & vbCrLf
        tmpStr &= "color: #777777;" & vbCrLf
        tmpStr &= "text-align: center;" & vbCrLf
        tmpStr &= "line-height: 21px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "a {" & vbCrLf
        tmpStr &= "color: #676767;" & vbCrLf
        tmpStr &= "text-decoration: none !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".pull-left {" & vbCrLf
        tmpStr &= "text-align: left;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".pull-right {" & vbCrLf
        tmpStr &= "text-align: right;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".header-lg," & vbCrLf
        tmpStr &= ".header-md," & vbCrLf
        tmpStr &= ".header-sm {" & vbCrLf
        tmpStr &= "font-size: 32px;" & vbCrLf
        tmpStr &= "font-weight: 700;" & vbCrLf
        tmpStr &= "line-height: normal;" & vbCrLf
        tmpStr &= "padding: 35px 0 0;" & vbCrLf
        tmpStr &= "color: #4d4d4d;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".header-md {" & vbCrLf
        tmpStr &= "font-size: 24px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".header-sm {" & vbCrLf
        tmpStr &= "padding: 5px 0;" & vbCrLf
        tmpStr &= "font-size: 18px;" & vbCrLf
        tmpStr &= "line-height: 1.3;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".content-padding {" & vbCrLf
        tmpStr &= "padding: 20px 0 5px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".mobile-header-padding-right {" & vbCrLf
        tmpStr &= "width: 290px;" & vbCrLf
        tmpStr &= "text-align: right;" & vbCrLf
        tmpStr &= "padding-left: 10px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".mobile-header-padding-left {" & vbCrLf
        tmpStr &= "width: 290px;" & vbCrLf
        tmpStr &= "text-align: left;" & vbCrLf
        tmpStr &= "padding-left: 10px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".free-text {" & vbCrLf
        tmpStr &= "width: 100% !important;" & vbCrLf
        tmpStr &= "padding: 10px 60px 0px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".button {" & vbCrLf
        tmpStr &= "padding: 30px 0;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".mini-block {" & vbCrLf
        tmpStr &= "border: 1px solid #e5e5e5;" & vbCrLf
        tmpStr &= "border-radius: 5px;" & vbCrLf
        tmpStr &= "background-color: #ffffff;" & vbCrLf
        tmpStr &= "padding: 12px 15px 15px;" & vbCrLf
        tmpStr &= "text-align: left;" & vbCrLf
        tmpStr &= "width: 253px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".mini-container-left {" & vbCrLf
        tmpStr &= "width: 278px;" & vbCrLf
        tmpStr &= "padding: 10px 0 10px 15px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".mini-container-right {" & vbCrLf
        tmpStr &= "width: 278px;" & vbCrLf
        tmpStr &= "padding: 10px 14px 10px 15px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".product {" & vbCrLf
        tmpStr &= "text-align: left;" & vbCrLf
        tmpStr &= "vertical-align: top;" & vbCrLf
        tmpStr &= "width: 175px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".total-space {" & vbCrLf
        tmpStr &= "padding-bottom: 8px;" & vbCrLf
        tmpStr &= "display: inline-block;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".item-table {" & vbCrLf
        tmpStr &= "padding: 50px 20px;" & vbCrLf
        tmpStr &= "width: 560px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".item {" & vbCrLf
        tmpStr &= "width: 300px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".mobile-hide-img {" & vbCrLf
        tmpStr &= "text-align: left;" & vbCrLf
        tmpStr &= "width: 125px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".mobile-hide-img img {" & vbCrLf
        tmpStr &= "border: 1px solid #e6e6e6;" & vbCrLf
        tmpStr &= "border-radius: 4px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".title-dark {" & vbCrLf
        tmpStr &= "text-align: left;" & vbCrLf
        tmpStr &= "border-bottom: 1px solid #cccccc;" & vbCrLf
        tmpStr &= "color: #4d4d4d;" & vbCrLf
        tmpStr &= "font-weight: 700;" & vbCrLf
        tmpStr &= "padding-bottom: 5px;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".item-col {" & vbCrLf
        tmpStr &= "padding-top: 20px;" & vbCrLf
        tmpStr &= "text-align: left;" & vbCrLf
        tmpStr &= "vertical-align: top;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= ".force-width-gmail {" & vbCrLf
        tmpStr &= "min-width:600px;" & vbCrLf
        tmpStr &= "height: 0px !important;" & vbCrLf
        tmpStr &= "line-height: 1px !important;" & vbCrLf
        tmpStr &= "font-size: 1px !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "</style>" & vbCrLf
        tmpStr &= "<style type=""text/css"" media=""screen"">" & vbCrLf
        tmpStr &= "@import url(http://fonts.googleapis.com/css?family=Oxygen:400,700);" & vbCrLf
        tmpStr &= "</style>" & vbCrLf
        tmpStr &= "<style type=""text/css"" media=""screen"">" & vbCrLf
        tmpStr &= "@media screen {" & vbCrLf
        tmpStr &= "/* Thanks Outlook 2013! */" & vbCrLf
        tmpStr &= "* {" & vbCrLf
        tmpStr &= "font-family: 'Oxygen', 'Helvetica Neue', 'Arial', 'sans-serif' !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "</style>" & vbCrLf
        tmpStr &= "<style type=""text/css"" media=""only screen and (max-width: 480px)"">" & vbCrLf
        tmpStr &= "/* Mobile styles */" & vbCrLf
        tmpStr &= "@media only screen and (max-width: 480px) {" & vbCrLf
        tmpStr &= "table[class*=""container-for-gmail-android""] {" & vbCrLf
        tmpStr &= "min-width: 290px !important;" & vbCrLf
        tmpStr &= "width: 100% !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "img[class=""force-width-gmail""] {" & vbCrLf
        tmpStr &= "display: none !important;" & vbCrLf
        tmpStr &= "width: 0 !important;" & vbCrLf
        tmpStr &= "height: 0 !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "table[class=""w320""] {" & vbCrLf
        tmpStr &= "width: 320px !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class*=""mobile-header-padding-left""] {" & vbCrLf
        tmpStr &= "width: 160px !important;" & vbCrLf
        tmpStr &= "padding-left: 0 !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class*=""mobile-header-padding-right""] {" & vbCrLf
        tmpStr &= "width: 160px !important;" & vbCrLf
        tmpStr &= "padding-right: 0 !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class=""header-lg""] {" & vbCrLf
        tmpStr &= "font-size: 24px !important;" & vbCrLf
        tmpStr &= "padding-bottom: 5px !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class=""content-padding""] {" & vbCrLf
        tmpStr &= "padding: 5px 0 5px !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class=""button""] {" & vbCrLf
        tmpStr &= "padding: 5px 5px 30px !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class*=""free-text""] {" & vbCrLf
        tmpStr &= "padding: 10px 18px 30px !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class~=""mobile-hide-img""] {" & vbCrLf
        tmpStr &= "display: none !important;" & vbCrLf
        tmpStr &= "height: 0 !important;" & vbCrLf
        tmpStr &= "width: 0 !important;" & vbCrLf
        tmpStr &= "line-height: 0 !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class~=""item""] {" & vbCrLf
        tmpStr &= "width: 140px !important;" & vbCrLf
        tmpStr &= "vertical-align: top !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class~=""quantity""] {" & vbCrLf
        tmpStr &= "width: 50px !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class~=""price""] {" & vbCrLf
        tmpStr &= "width: 90px !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class=""item-table""] {" & vbCrLf
        tmpStr &= "padding: 30px 20px !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "td[class=""mini-container-left""]," & vbCrLf
        tmpStr &= "td[class=""mini-container-right""] {" & vbCrLf
        tmpStr &= "padding: 0 15px 15px !important;" & vbCrLf
        tmpStr &= "display: block !important;" & vbCrLf
        tmpStr &= "width: 290px !important;" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "}" & vbCrLf
        tmpStr &= "</style>" & vbCrLf
        tmpStr &= "</head>" & vbCrLf
        tmpStr &= "<body bgcolor=""#f7f7f7"">" & vbCrLf
        tmpStr &= "<table align=""center"" cellpadding=""0"" cellspacing=""0"" class=""container-for-gmail-android"" width=""100%"">" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td align=""left"" valign=""top"" width=""100%"" style=""background:repeat-x url(http://s3.amazonaws.com/swu-filepicker/4E687TRe69Ld95IDWyEg_bg_top_02.jpg) #ffffff;"">" & vbCrLf
        tmpStr &= "<center>" & vbCrLf
        tmpStr &= "<img src=""http://s3.amazonaws.com/swu-filepicker/SBb2fQPrQ5ezxmqUTgCr_transparent.png"" class=""force-width-gmail"">" & vbCrLf
        tmpStr &= "<table cellspacing=""0"" cellpadding=""0"" width=""100%"" bgcolor=""#ffffff"" background=""http://s3.amazonaws.com/swu-filepicker/4E687TRe69Ld95IDWyEg_bg_top_02.jpg"" style=""background-color:transparent"">" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td width=""100%"" height=""80"" valign=""top"" style=""text-align: center; vertical-align:middle;"">" & vbCrLf
        tmpStr &= "<!--[if gte mso 9]>"
        tmpStr &= "<v:rect xmlns:v=""urn:schemas-microsoft-com:vml"" fill=""true"" stroke=""false"" style=""mso-width-percent:1000;height:80px; v-text-anchor:middle;"">" & vbCrLf
        tmpStr &= "<v:fill type=""tile"" src=""http://s3.amazonaws.com/swu-filepicker/4E687TRe69Ld95IDWyEg_bg_top_02.jpg"" color=""#ffffff"" />" & vbCrLf
        tmpStr &= "<v:textbox inset=""0,0,0,0"">" & vbCrLf
        tmpStr &= "<![endif]-->" & vbCrLf
        tmpStr &= "<center>" & vbCrLf
        tmpStr &= "<table cellpadding=""0"" cellspacing=""0"" width=""600"" class=""w320"">" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td class=""pull-left mobile-header-padding-left"" style=""vertical-align: middle;"">" & vbCrLf
        tmpStr &= "<a href=""https://camarabolivar.com.ar""><img width=""231"" height=""47"" src=""https://camarabolivar.com.ar/img/2019/camara_bolivar_logo.jpg"" alt=""Cámara Bolívar""></a>" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "<td class=""pull-right mobile-header-padding-right"" style=""color: #4d4d4d;"">" & vbCrLf
        tmpStr &= "<a href=""https://camarabolivar.com.ar/""><img style=""margin-right: 10px"" width=""20"" height=""20"" src=""https://camarabolivar.com.ar/img/2019/w.png"" alt=""web"" /></a>" & vbCrLf
        tmpStr &= "<a href=""https://twitter.com/camara_bolivar""><img style=""margin-right: 10px"" width=""20"" height=""20"" src=""https://camarabolivar.com.ar/img/2019/t.png"" alt=""twitter"" /></a>" & vbCrLf
        tmpStr &= "<a href=""https://www.facebook.com/camarabolivar/""><img style=""margin-right: 10px"" width=""20"" height=""20"" src=""https://camarabolivar.com.ar/img/2019/f.png"" alt=""facebook"" /></a>" & vbCrLf
        tmpStr &= "<a href=""https://www.instagram.com/camarabolivar/""><img style=""margin-right: 10px"" width=""20"" height=""20"" src=""https://camarabolivar.com.ar/img/2019/i.png"" alt=""Instagram"" /></a>" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "</table>" & vbCrLf
        tmpStr &= "</center>" & vbCrLf
        tmpStr &= "<!--[if gte mso 9]>"
        tmpStr &= "</v:textbox>" & vbCrLf
        tmpStr &= "</v:rect>" & vbCrLf
        tmpStr &= "<![endif]-->" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "</table>" & vbCrLf
        tmpStr &= "</center>" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td align=""center"" valign=""top"" width=""100%"" style=""background-color: #f7f7f7;"" class=""content-padding"">" & vbCrLf
        tmpStr &= "<center>" & vbCrLf
        tmpStr &= "<table cellspacing=""0"" cellpadding=""0"" width=""600"" class=""w320"">" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td class=""header-lg"">" & vbCrLf
        tmpStr &= "¡Te acercamos tu <br /><span>Factura Digital</span>!" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td class=""free-text"">" & vbCrLf
        tmpStr &= $"<br/><span style=""font-size:18px"">Hola <strong>{fex.RazonSocialReceptor}</strong>,</span><br/>" & vbCrLf
        tmpStr &= "<br/>{#}<br/><br/>" & vbCrLf
        tmpStr &= "Con este correo va adjunta una factura eléctronica con el siguiente detalle:" & vbCrLf
        tmpStr &= "<br /><br />" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "</table>" & vbCrLf
        tmpStr &= "</center>" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td align=""center"" valign=""top"" width=""100%"" style=""background-color: #ffffff;  border-top: 1px solid #e5e5e5; border-bottom: 1px solid #e5e5e5;"">" & vbCrLf
        tmpStr &= "<center>" & vbCrLf
        tmpStr &= "<table cellpadding=""0"" cellspacing=""0"" width=""600"" class=""w320"">" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td class=""item-table"">" & vbCrLf
        tmpStr &= "<table cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td class=""title-dark"" width=""300"">" & vbCrLf
        tmpStr &= "Detalle" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "<td class=""title-dark"" width=""163"">" & vbCrLf
        tmpStr &= "Cantidad" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "<td class=""title-dark"" width=""97"">" & vbCrLf
        tmpStr &= "Subtotal" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf

        For Each detalle As AfipFacturaDetalle In fe.Detalles
            tmpStr &= "<tr>" & vbCrLf
            tmpStr &= "<td class=""item-col item"">" & vbCrLf
            tmpStr &= "<table cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbCrLf
            tmpStr &= "<tr>" & vbCrLf
            tmpStr &= "<td class=""product"">" & vbCrLf
            tmpStr &= $"<span style=""color: #4d4d4d; font-weight:bold;"">{detalle.ProductoServicio}</span> <br />" & vbCrLf
            tmpStr &= "&nbsp;" & vbCrLf
            tmpStr &= "</td>" & vbCrLf
            tmpStr &= "</tr>" & vbCrLf
            tmpStr &= "</table>" & vbCrLf
            tmpStr &= "</td>" & vbCrLf
            tmpStr &= "<td class=""item-col quantity"">" & vbCrLf
            tmpStr &= $"{detalle.Cantidad}" & vbCrLf
            tmpStr &= "</td>" & vbCrLf
            tmpStr &= "<td class=""item-col"">" & vbCrLf
            tmpStr &= $"{Utils.ToMoneyFormat(detalle.PrecioUnitario - ((detalle.PrecioUnitario * detalle.BonificacionPercent) / 100))}" & vbCrLf
            tmpStr &= "</td>" & vbCrLf
            tmpStr &= "</tr>" & vbCrLf
        Next

        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td class=""item-col item mobile-row-padding""></td>" & vbCrLf
        tmpStr &= "<td class=""item-col quantity""></td>" & vbCrLf
        tmpStr &= "<td class=""item-col price""></td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td class=""item-col item"">" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "<td class=""item-col quantity"" style=""text-align:right; padding-right: 10px; border-top: 1px solid #cccccc;"">" & vbCrLf
        tmpStr &= "<span class=""total-space"" style=""font-weight: bold; color: #4d4d4d"">Total</span>" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "<td class=""item-col price"" style=""text-align: left; border-top: 1px solid #cccccc;"">" & vbCrLf
        tmpStr &= $"<span class=""total-space"" style=""font-weight:bold; color: #4d4d4d"">{Utils.ToMoneyFormat(fe.ImporteTotal)}</span>" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "</table>" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "</table>" & vbCrLf
        tmpStr &= "</center>" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td align=""center"" valign=""top"" width=""100%"" style=""background-color: #f7f7f7; height: 100px;"">" & vbCrLf
        tmpStr &= "<center>" & vbCrLf
        tmpStr &= "<table cellspacing=""0"" cellpadding=""0"" width=""600"" class=""w320"">" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td style=""padding: 25px 0 25px;font-size:12px;"">" & vbCrLf
        tmpStr &= "<strong>Cámara Bolívar</strong><br />" & vbCrLf
        tmpStr &= "Comercio, Industria y Servicios<br /><br />" & vbCrLf
        tmpStr &= "Las Heras 45 - San Carlos de Bolívar<br />" & vbCrLf
        tmpStr &= "Buenos Aires - Argentina<br />" & vbCrLf
        tmpStr &= "(02314) 42-7327<br />" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "<tr>" & vbCrLf
        tmpStr &= "<td style=""padding: 25px 0 25px"">" & vbCrLf
        tmpStr &= "<small style=""font-size:10px"">" & vbCrLf
        tmpStr &= "DERECHOS SOBRE SUS DATOS PERSONALES. Le recordamos que, con el fin de intentar resguardar su seguridad, en los e-mails que Cámara Comercial e Industrial de Bolívar le envíe no solicitará ningún tipo de ingreso de datos ni incluirá en los mismos links directos a páginas en internet que se los soliciten. Por ello, le sugerimos nunca proveer información vía e-mail o accediendo a links directos contenidos en e-mails, incluso cuando la información fuera aparentemente solicitada por Cámara Comercial e Industrial de Bolívar. Dirección Nacional de Protección de Datos Personales. Ley n° 25.326 art. 27 (archivos, registros o bancos de datos con fines de publicidad). Inc. 3. El titular podrá en cualquier momento solicitar el retiro o bloqueo de su nombre de los bancos de datos a los que se refiere el presente artículo. Dto. N° 1558/2001 art. 27. En toda comunicación con fines de publicidad que se realice por correo, teléfono, correo electrónico, internet u otro medio a distancia a conocer, se deberá indicar, en forma expresa y destacada, la posibilidad del titular del dato de solicitar el retiro o bloqueo, total o parcial, de su nombre de la base de datos. A pedido del interesado, se deberá informar el nombre del responsable o usuario del banco de datos que proveyó la información. El titular de los datos personales tiene la facultad de ejercer el derecho de acceso a los mismos en forma gratuita a intervalos no inferiores a seis meses, salvo que se acredite un interés legítimo al efecto conforme lo establecido en el artículo 14, inciso 3 de la ley N° 25.326. La Dirección Nacional de Protección de Datos Personales, Órgano de Control de la Ley N° 25.326, tiene la atribución de atender las denuncias y reclamos que se interpongan con relación al incumplimiento de las normas sobre Protección de Datos Personales. Para contactar a la Dirección Nacional de Protección de Datos Personales podrá dirigirse a Sarmiento 1118, piso 5° de la Ciudad Autónoma de Buenos Aires (C1041AAX) o por teléfono al 5300-4000 Int. 76706/24/23/42. </small><br>" & vbCrLf
        tmpStr &= "<br><small>Desarrollo <b><a href=""https://logico.com.ar"">⬢ Lógico</a><b></small>" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "</table>" & vbCrLf
        tmpStr &= "</center>" & vbCrLf
        tmpStr &= "</td>" & vbCrLf
        tmpStr &= "</tr>" & vbCrLf
        tmpStr &= "</table>" & vbCrLf
        tmpStr &= "</div>" & vbCrLf
        tmpStr &= "</body>" & vbCrLf
        tmpStr &= "</html>" & vbCrLf
        Return tmpStr
    End Function
End Module
