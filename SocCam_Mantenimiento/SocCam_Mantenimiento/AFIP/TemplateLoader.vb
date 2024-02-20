Imports System.IO
Imports System.Xml

Public Class TemplateLoader

    Public Shared Sub LoadTemplate(doc As XmlDocument, file As String)
        doc.LoadXml(GetAuthXML)
    End Sub

    Public Shared Function GetAuthXML() As String
        Dim tmp As String
        tmp &= "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbCrLf
        tmp &= "<loginTicketRequest>" & vbCrLf
        tmp &= "  <header>" & vbCrLf
        tmp &= "    <uniqueId></uniqueId>" & vbCrLf
        tmp &= "    <generationTime></generationTime>" & vbCrLf
        tmp &= "    <expirationTime></expirationTime>" & vbCrLf
        tmp &= "  </header>" & vbCrLf
        tmp &= "  <service></service>" & vbCrLf
        tmp &= "</loginTicketRequest>" & vbCrLf
        Return tmp
    End Function


End Class
