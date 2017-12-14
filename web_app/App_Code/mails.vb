Imports System.IO
Imports System.Data.OleDb
Imports commonLib

Public Class MailTemplate

    Dim commonUtil As New commonLib.Utils
    Dim sysConfiguration As New SystemConfiguration

    Private Function PopulateBody(ByRef reader As StreamReader, ByRef mailData As Dictionary(Of String, String)) As String
        Dim urlHost As String = String.Empty
        Dim body As String = String.Empty
        body = reader.ReadToEnd
        For Each kvp As KeyValuePair(Of String, String) In mailData
            If kvp.Key.Substring(0, 1) = "{" Then
                body = body.Replace(kvp.Key, kvp.Value)
            End If
        Next kvp

        'Replace image headers
        urlHost = sysConfiguration.getSystemUrl()
        body = body.Replace("{sap_logo}", urlHost + "images/email-templates/logo.png")
        body = body.Replace("{sap_header}", urlHost + "images/email-templates/header.jpg")

        'Return replaced body
        Return body
    End Function


    Private Sub SendHtmlFormattedEmail(ByVal recipientEmail As String, ByVal subject As String, ByVal body As String, ByVal ai_Id As String)

        Dim dbconn As OleDbConnection
        Dim dbcomm As OleDbCommand
        Dim sql As String
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New Logging

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        If (String.IsNullOrEmpty(ai_Id)) Then
            sql = "INSERT INTO send (recipients, subject, body) VALUES ('" + recipientEmail + "', '" + subject.Replace("&#34;", """").Replace("'", "&#39;") + "', '" + body.Replace("&#34;", """").Replace("'", "&#39;") + "')"
        Else
            sql = "INSERT INTO send (recipients, subject, body, ai_id) VALUES ('" + recipientEmail + "', '" + subject.Replace("&#34;", """").Replace("'", "&#39;") + "', '" + body.Replace("&#34;", """").Replace("'", "&#39;") + "', " + ai_Id + ")"
        End If
        'sql = "INSERT INTO send (recipients, subject, body) VALUES ('" + recipientEmail + "', '" + subject.Replace("&#34;", """").Replace("'", "&#39;") + "', '" + body.Replace("&#34;", """").Replace("'", "&#39;") + "')"

        dbcomm = New OleDbCommand(sql, dbconn)
        dbcomm.ExecuteNonQuery()
        dbcomm.CommandText = "SELECT @@IDENTITY"
        dbcomm.ExecuteScalar()

        dbconn.Close()

    End Sub

    Public Sub SendNotificationMail(ByRef mailData As Dictionary(Of String, String))
        Dim mailTemplate As String = ""
        Dim mailSubject As String = ""
        Dim mailBody As String = ""
        Dim ai_Id As String = ""

        Dim mailRecepient As String = mailData("to")
        If mailData.ContainsKey("{ai_id}") AndAlso Not String.IsNullOrEmpty(mailData("{ai_id}")) Then
            ai_Id = mailData("{ai_id}")
        End If

        mailTemplate = commonUtil.getEmailTemplate(mailData("mail"))

        'GET EMAIL SUBJECT
        'When the subject is provided in the map
        If mailData.ContainsKey("{subject}") AndAlso Not String.IsNullOrEmpty(mailData("{subject}")) Then
            mailSubject = mailData("{subject}")
        Else
            'Get default email subject
            mailSubject = commonUtil.getDefaultEmailSubject(mailData("mail"))
        End If

        Dim reader As StreamReader = New StreamReader(HttpContext.Current.Server.MapPath(mailTemplate))

        mailBody = PopulateBody(reader, mailData)

        SendHtmlFormattedEmail(mailRecepient, mailSubject, mailBody, ai_Id)

    End Sub

End Class