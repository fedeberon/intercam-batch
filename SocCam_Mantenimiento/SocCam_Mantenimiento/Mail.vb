Imports helix
Imports System.Net.Mail
Imports System.Net.Mail.Attachment
Imports System.Text.RegularExpressions
Imports System.ComponentModel

Public Class Mail

    Private _body As String = ""

    Public Property Body As String
        Set(value As String)
            _body = value
        End Set
        Get
            Return _body
        End Get
    End Property

    Public Property CSSCode As String = ""
    Public Property HTMLCode As String = ""
    Public Property Subject As String = ""
    Public Property IsHTML As Boolean = True
    Public Property ToAddress As String = ""
    Public Property FromAddress As String = ""
    Public Property FromName As String = ""
    Public Property ReplyAddress As String = ""
    Public Property Smtp_username As String = ""
    Public Property Smtp_password As String = ""
    Public Property Smtp_host As String = ""
    Public Property Smtp_port As Integer = 587
    Public Property Smtp_SSL As Boolean = True
    Public Property Adjunto As String = ""

    Public Property Bcc As New List(Of String)

    Private Const DEFAULT_CSS_PLAIN_TEXT As String = "<style>body{padding:0;margin:0;font-family:""Courier New"";}</style>"
    Private _htmlPreCode As String = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org=/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><meta name=""viewport"" content=""width=device-width, initial-scale=1.0""/><default-css></default-css><style></style></head><body>"

    Private _htmlBody As String = "<h1>Previsualización</h1>"

    Private _htmlPostCode As String = "</body></html>"

    Public Function LoadPlantilla(ByVal plantillaId As Integer, ByVal sqle As SQLEngine) As Boolean
        Dim plantilla As New MailTemplate
        plantilla.sqle = sqle
        If plantilla.LoadMe(plantillaId) Then
            Me.CSSCode = plantilla.css
            Me.HTMLCode = plantilla.body
            Me.IsHTML = plantilla.esHTML
            Return True
        End If
        Return False
    End Function

    Private Shared Sub SendCompletedCallback(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs)
        ' Get the unique identifier for this asynchronous operation.
        Dim token As String = CStr(e.UserState)

        If e.Cancelled Then
            Console.WriteLine("[{0}] Send canceled.", token)
        End If
        If e.Error IsNot Nothing Then
            Console.WriteLine("[{0}] {1}", token, e.Error.ToString())
        Else
            Console.WriteLine("Message sent.")
        End If
    End Sub


    Public Function SendMail(Optional ByVal isHTML5 As Boolean = False, Optional Cco As List(Of String) = Nothing, Optional simulation As Boolean = False, Optional async As Boolean = False) As Boolean
        Try
            Dim SmtpServer As New SmtpClient()
            Dim mail As New MailMessage()
            SmtpServer.Credentials = New Net.NetworkCredential(Smtp_username, Smtp_password)
            SmtpServer.Host = Smtp_host
            SmtpServer.Port = Smtp_port
            SmtpServer.EnableSsl = Smtp_SSL
            SmtpServer.Timeout = 180000

            mail = New MailMessage()

            ' CUIDADO: puede que el nombre de usuario no sea el mail
            mail.From = New MailAddress(Smtp_username, FromName)

            If ReplyAddress <> "" Then
                mail.ReplyToList.Add(ReplyAddress)
            End If

            ' Carga la lista de copias ocultas
            For Each tmpStr In Bcc
                mail.Bcc.Add(tmpStr)
            Next


            mail.Subject = Subject

            mail.Body = SanitizeTemplate(Compose(isHTML5))

            mail.IsBodyHtml = IsHTML

            If ToAddress <> "" Then
                mail.To.Add(ToAddress)
            End If


            If Adjunto.Length <> 0 Then
                Dim data As Attachment = New Attachment(Adjunto)
                mail.Attachments.Add(data)
            End If


            If Not simulation Then
                If async Then
                    AddHandler SmtpServer.SendCompleted, AddressOf SendCompletedCallback
                    Dim userState As String = "Sending message..."
                    SmtpServer.SendAsync(mail, userState)
                Else
                    SmtpServer.Send(mail)
                End If
            End If

            SmtpServer.Dispose()


            Return True
        Catch ex As Exception
            Debug.Print(ex.ToString)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Compone el cuerpo del mail
    ''' </summary>
    ''' <returns>El cuerpo del mensaje</returns>
    ''' <remarks></remarks>
    Public Function Compose(Optional ByVal isHTML5 As Boolean = False) As String
        Dim tmpMail As String = ""
        If IsHTML Then
            If isHTML5 Then
                tmpMail = HTMLCode
            Else
                tmpMail = _htmlPreCode.Replace("<style></style>", "<style>" & CSSCode & "</style>") & HTMLCode & _htmlPostCode
            End If

        Else
            tmpMail = _htmlPreCode.Replace("<style></style>", DEFAULT_CSS_PLAIN_TEXT) & "<pre><code>" & HTMLCode.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;") & "</code></pre>" & _htmlPostCode
        End If

        Return tmpMail
    End Function

    ''' <summary>
    ''' Limpia todo rastro de las variables.
    ''' </summary>
    ''' <remarks>Limpia todas las variables que no son utilizadas o que hayan quedadas mal escritas</remarks>
    Public Function SanitizeTemplate(ByVal htmlCode As String) As String
        Dim firstIndex As Integer = 0
        Dim lastIndex As Integer = 0
        Dim counter As Byte = 0

        Dim tmpHTML As String = htmlCode


        ' #BUG
        Return tmpHTML

        While (tmpHTML.IndexOf("{{") <> -1) And (tmpHTML.IndexOf("}}") <> -1) And (counter < 250)                 ' En el caso que haya un fallo, agrego un contador de control (poco probable que usen 250 variables en una misma plantilla)
            firstIndex = tmpHTML.IndexOf("{{")
            lastIndex = tmpHTML.IndexOf("}}")

            tmpHTML = tmpHTML.Remove(firstIndex, (lastIndex - firstIndex) + 2)  ' + 2 para compensar la base 0 y el segundo }
            counter += 1
        End While

        Return tmpHTML
    End Function

    ''' <summary>
    ''' Reemplaza todas las ocurrencias de la variable por el valor
    ''' </summary>
    ''' <param name="htmlCode">Codigo donde se encuentran las variables</param>
    ''' <param name="varName">Variable a reemplazar</param>
    ''' <param name="varValue">Valor de la variable</param>
    ''' <returns>Cadena con los valores en lugar de la variable</returns>
    ''' <remarks></remarks>
    Public Function ReplaceVariable(ByVal htmlCode As String, ByVal varName As String, ByVal varValue As String) As String
        Return htmlCode.Replace(varName, varValue)
    End Function

    Public Function IsValidMailAddress(ByVal emailAddress As String) As Boolean
        emailAddress = emailAddress.ToLower
        If emailAddress.Contains("ñ") Then Return False
        If emailAddress.Contains("á") Then Return False
        If emailAddress.Contains("é") Then Return False
        If emailAddress.Contains("í") Then Return False
        If emailAddress.Contains("ó") Then Return False
        If emailAddress.Contains("ú") Then Return False

        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            IsValidMailAddress = True
        Else
            IsValidMailAddress = False
        End If
    End Function

End Class
