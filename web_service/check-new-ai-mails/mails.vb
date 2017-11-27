Imports System.IO
Imports System.Data.OleDb
Imports System.Web.HttpUtility
Imports System.Text
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports MailKit
Imports MailKit.Security
Imports System.Text.RegularExpressions
Imports commonLib


'Highligth COLOR
'#FFFC6E

Public Class MailTemplate

    Dim Smtp_Server As New SmtpClient
    Dim e_mail As New MailMessage
    Dim systemConfig As New SystemConfiguration
    Dim appConfiguration As New AppSettings
    Dim utils As New Utils

    'DEFINE GLOBAL VARIABLES
    Public currentEnv As String = "production"
    Public imapServer As String = "imap.global.corp.sap"
    Public imapPort As Integer = 993
    Public isSSL = True
    Public hostSMTP As String = "mail.sap.corp"
    Public emailUser As String = "asa1_sap_mktg_in_ac@global.corp.sap\sap_marketing_in_action"
    Public emailPass As String = "BAoR}:qKQSkzSBO'#4pQ"
    Public emailAddressFrom As String = "sap_marketing_in_action@sap.com"
    Dim syscfg As New SysConfig

    Function ResolveDisplayNameToSMTP(ByVal userName As String) As String 'ByRef session As Outlook.NameSpace,
        Dim users As New SapUser

        Return users.getMailByName(userName)

    End Function

    Function smtpList(ByVal recip As String) As List(Of String)
        Dim result As New List(Of String)

        If Not String.IsNullOrEmpty(recip) Then

            Dim mailToList() As String = recip.Split(";")

            For Each address In mailToList
                If Not String.IsNullOrEmpty(address) Then
                    result.Add(ResolveDisplayNameToSMTP(address))
                End If
            Next

        End If

        Return result

    End Function

    Function exToSMTP(ByVal recip As String) As String
        Dim result As String = ""

        If Not String.IsNullOrEmpty(recip) Then

            Dim mailToList() As String = recip.Split(";")

            For Each address In mailToList
                If Not String.IsNullOrEmpty(address) Then
                    result = result & ResolveDisplayNameToSMTP(address) & ";"
                End If
            Next

        End If

        Return result

    End Function

    Private Function StripHTML(ByVal source As String) As String
        Try
            Dim result As String

            ' Remove HTML Development formatting
            ' Replace line breaks with space
            ' because browsers inserts space
            result = source.Replace(vbCr, " ")
            ' Replace line breaks with space
            ' because browsers inserts space
            result = result.Replace(vbLf, " ")
            ' Remove step-formatting
            result = result.Replace(vbTab, String.Empty)
            ' Remove repeating spaces because browsers ignore them
            result = RegularExpressions.Regex.Replace(result, "( )+", " ")

            ' Remove the header (prepare first by clearing attributes)
            result = RegularExpressions.Regex.Replace(result, "<( )*head([^>])*>", "<head>", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*head( )*>)", "</head>", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "(<head>).*(</head>)", String.Empty, RegularExpressions.RegexOptions.IgnoreCase)

            ' remove all scripts (prepare first by clearing attributes)
            result = RegularExpressions.Regex.Replace(result, "<( )*script([^>])*>", "<script>", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*script( )*>)", "</script>", RegularExpressions.RegexOptions.IgnoreCase)
            'result = System.Text.RegularExpressions.Regex.Replace(result,
            '         @"(<script>)([^(<script>\.</script>)])*(</script>)",
            '         string.Empty,
            '         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            result = RegularExpressions.Regex.Replace(result, "(<script>).*(</script>)", String.Empty, RegularExpressions.RegexOptions.IgnoreCase)

            ' remove all styles (prepare first by clearing attributes)
            result = RegularExpressions.Regex.Replace(result, "<( )*style([^>])*>", "<style>", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*style( )*>)", "</style>", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "(<style>).*(</style>)", String.Empty, RegularExpressions.RegexOptions.IgnoreCase)

            ' insert tabs in spaces of <td> tags
            result = RegularExpressions.Regex.Replace(result, "<( )*td([^>])*>", vbTab, RegularExpressions.RegexOptions.IgnoreCase)

            ' insert line breaks in places of <BR> and <LI> tags
            result = RegularExpressions.Regex.Replace(result, "<( )*br( )*>", vbCr, RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "<( )*li( )*>", vbCr, RegularExpressions.RegexOptions.IgnoreCase)

            ' insert line paragraphs (double line breaks) in place
            ' if <P>, <DIV> and <TR> tags
            result = RegularExpressions.Regex.Replace(result, "<( )*div([^>])*>", vbCr & vbCr, RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "<( )*tr([^>])*>", vbCr & vbCr, RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "<( )*p([^>])*>", vbCr & vbCr, RegularExpressions.RegexOptions.IgnoreCase)

            ' Remove remaining tags like <a>, links, images,
            ' comments etc - anything that's enclosed inside < >
            result = RegularExpressions.Regex.Replace(result, "<[^>]*>", String.Empty, RegularExpressions.RegexOptions.IgnoreCase)

            ' replace special characters:
            result = RegularExpressions.Regex.Replace(result, " ", " ", RegularExpressions.RegexOptions.IgnoreCase)

            result = RegularExpressions.Regex.Replace(result, "&bull;", " * ", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "&lsaquo;", "<", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "&rsaquo;", ">", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "&trade;", "(tm)", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "&frasl;", "/", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "&lt;", "<", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "&gt;", ">", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "&copy;", "(c)", RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "&reg;", "(r)", RegularExpressions.RegexOptions.IgnoreCase)
            ' Remove all others. More can be added, see
            ' http://hotwired.lycos.com/webmonkey/reference/special_characters/
            result = RegularExpressions.Regex.Replace(result, "&(.{2,6});", String.Empty, RegularExpressions.RegexOptions.IgnoreCase)

            ' for testing
            'System.Text.RegularExpressions.Regex.Replace(result,
            '       this.txtRegex.Text,string.Empty,
            '       System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            ' make line breaking consistent
            result = result.Replace(vbLf, vbCr)

            ' Remove extra line breaks and tabs:
            ' replace over 2 breaks with 2 and over 4 tabs with 4.
            ' Prepare first to remove any whitespaces in between
            ' the escaped characters and remove redundant tabs in between line breaks
            result = RegularExpressions.Regex.Replace(result, "(" & vbCr & ")( )+(" & vbCr & ")", vbCr & vbCr, RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "(" & vbTab & ")( )+(" & vbTab & ")", vbTab & vbTab, RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "(" & vbTab & ")( )+(" & vbCr & ")", vbTab & vbCr, RegularExpressions.RegexOptions.IgnoreCase)
            result = RegularExpressions.Regex.Replace(result, "(" & vbCr & ")( )+(" & vbTab & ")", vbCr & vbTab, RegularExpressions.RegexOptions.IgnoreCase)
            ' Remove redundant tabs
            result = RegularExpressions.Regex.Replace(result, "(" & vbCr & ")(" & vbTab & ")+(" & vbCr & ")", vbCr & vbCr, RegularExpressions.RegexOptions.IgnoreCase)
            ' Remove multiple tabs following a line break with just one tab
            result = RegularExpressions.Regex.Replace(result, "(" & vbCr & ")(" & vbTab & ")+", vbCr & vbTab, RegularExpressions.RegexOptions.IgnoreCase)
            ' Initial replacement target string for line breaks
            Dim breaks As String = vbCr & vbCr & vbCr
            ' Initial replacement target string for tabs
            Dim tabs As String = vbTab & vbTab & vbTab & vbTab & vbTab
            For index As Integer = 0 To result.Length - 1
                result = result.Replace(breaks, vbCr & vbCr)
                result = result.Replace(tabs, vbTab & vbTab & vbTab & vbTab)
                breaks = breaks & Convert.ToString(vbCr)
                tabs = tabs & Convert.ToString(vbTab)
            Next

            ' That's it.
            Return result
        Catch
            Return source
        End Try
    End Function

    Private Function dictMail(ByRef message As MimeKit.MimeMessage) As Dictionary(Of String, String)
        'CONVERT THE MAIL INTO STRINGS INSTEAD OF A EMAIL CLASS
        'AS GIFT RETURNS THE OWNERS LIST INTO THE DICTIONARY
        Dim mailToSystem As Boolean = False
        Dim mailsList As New List(Of String)
        Dim ownersList As New List(Of String)
        Dim ownersDBList As New List(Of String)
        Dim mailOwn As String
        Dim syscfg As New SysConfig

        Dim users As New SapUser
        Dim mailFile As String = ""

        mailFile = appConfiguration.emailStorePath & message.MessageId & ".eml"
        If Not File.Exists(mailFile) Then
            message.WriteTo(mailFile)
        End If

        If message.To.Count > 0 Then
            Dim emailNameList As List(Of String) = GetMailNames(message.To)
            For Each emailName As String In emailNameList
                If Not String.IsNullOrEmpty(emailName) Then
                    ownersList.AddRange(users.getOwnersEmailFromRecipient(emailName))
                End If
            Next
        End If

        If message.Cc.Count > 0 Then
            Dim emailNameList As List(Of String) = GetMailNames(message.Cc)
            For Each emailName As String In emailNameList
                If Not String.IsNullOrEmpty(emailName) Then
                    ownersList.AddRange(users.getOwnersEmailFromRecipient(emailName))
                End If
            Next
        End If

        Dim result As New Dictionary(Of String, String)
        Dim messageBody As String

        result.Add("from", ResolveDisplayNameToSMTP(JoinMailNames(message.From)))
        result.Add("subject", message.Subject)

        '#####################################
        '#####################################
        '#####################################
        messageBody = StripHTML(message.HtmlBody)
        messageBody.Replace("'", "''")
        log("dictMail - body: " & messageBody)

        If message.Body IsNot Nothing Then log("Message body is not null")
        result.Add("body", messageBody)
        result.Add("date", message.Date.Date)
        result.Add("id", message.MessageId)
        result.Add("raw", Encoding.ASCII.GetString(Encoding.UTF8.GetBytes(message.HtmlBody)).Replace("'", "''"))

        'ADD EACH OWNER: own1, own2, own3...
        Dim j As Integer = 1
        For Each mailOwn In ownersList
            'Filter email system from owners list
            If (Not isMailToSystem(mailOwn) And Not String.IsNullOrEmpty(mailOwn)) Then
                result.Add("own" + j.ToString, mailOwn)
                j = j + 1
            End If
        Next

        Return result

    End Function

    Public Function isMailToSystem(ByVal mailList As String) As Boolean
        Dim syscfg As New SystemConfiguration
        Dim result As Boolean = False

        Return mailList.IndexOf(syscfg.getSystemEmail) >= 0

    End Function

    Public Sub CheckMail()
        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim actions As New SapActions
        Dim dbDate As Date
        Dim log_dict As New Dictionary(Of String, String)
        Dim newLog As New Logging
        Dim lastMails As Collection = New Collection
        Dim c As Integer = 0
        Dim inbox As IMailFolder
        Dim imapClient As New Net.Imap.ImapClient

        log("CheckMail started")

        'GET LAST MAIL PROCESSED DATETIME AND ID
        dbDate = syscfg.getMailLastCheck

        'CONNECT MAPI ACCOUNT
        Using imapClient

            'Connect to server
            imapClient.ServerCertificateValidationCallback = AddressOf CertificateValidationCallBack
            imapClient.Connect(imapServer, imapPort, SecureSocketOptions.Auto)
            imapClient.Authenticate(emailUser, emailPass)

            'Select inbox as default folder to examine. It's always available on all IMAP servers
            inbox = imapClient.Inbox
            inbox.Open(FolderAccess.ReadWrite)

            log("Number of messages is: " & inbox.Count)
            log("Number of unseen messages is: " & inbox.Search(Search.SearchQuery.NotSeen).Count)
            log("Starting process to download and parse each message")
            log("last email processed as date: " & dbDate.Date)

            For Each uid As UniqueId In inbox.Search(Search.SearchQuery.NotSeen)
                Dim email = inbox.GetMessage(uid) ' Download and parse each message

                log("email date: " & email.Date.Date)
                If email.Date.Date >= dbDate.Date Then
                    lastMails.Add(email)
                    log("Adding:" & email.Subject)
                    log("EntryID: " & email.MessageId)
                    inbox.SetFlags(uid, MessageFlags.Seen, True, Nothing)
                Else
                    log("Old message - NOT adding:" & email.Subject)
                    log("EntryID: " & email.MessageId)
                    inbox.SetFlags(uid, MessageFlags.Deleted, True, Nothing)
                End If

            Next

            'clean old seen messages
            If (inbox.Search(Search.SearchQuery.Seen).Count > 0) Then
                For Each uid As UniqueId In inbox.Search(Search.SearchQuery.Seen)
                    Dim email = inbox.GetMessage(uid) ' Download and parse each message
                    If (email IsNot Nothing) Then
                        log("Email marked or deletion with entryID: " & email.MessageId & email.Subject)
                        inbox.SetFlags(uid, MessageFlags.Deleted, True, Nothing)
                    End If
                Next
            End If

            'force to purgue the inbox for all messages marked for deletion
            inbox.Expunge()

        End Using

        log("Last mail: " & dbDate.ToString("MM/dd/yyyy hh:mm:ss") & " Count: " & lastMails.Count.ToString)

        'CREATE TEXT FILE
        Dim sb As New StringBuilder()

        sb.AppendLine("<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>")
        sb.AppendLine("<html xmlns='http://www.w3.org/1999/xhtml'>")
        sb.AppendLine("<head><title>Inbox</title><link href='css/style.css' rel='stylesheet' type='text/css' /><script type='text/javascript' src='js/jquery-1.7.2.js'></script><script type='text/javascript' src='js/jquery-latest.js'></script><script type='text/javascript' src='js/jquery.tablesorter.js'></script><script type='text/javascript'>jQuery(function () {$(document).ready(function() {$('#inbox').tablesorter();} );});</script></head><body>")
        sb.AppendLine("<h3>Inbox Items<h3><br>" + dbDate.ToString("MM/dd/yyyy HH:mm") + " [" + lastMails.Count.ToString + "]<table id='inbox' class='tablesorter'><thead><tr><th class='header headerSortDown'>creation</th><th class='header'>type</th><th class='header'>due</th><th class='header'>from</th><th class='header'>subject</th><th class='header'>to</th><th class='header'>cc</th><th class='header'>body</th></tr></thead><tbody>")

        For Each message As MimeKit.MimeMessage In lastMails

            log("Processing: " & message.Subject)
            log("To:" & JoinMailNames(message.To) & " CC:" & JoinMailNames(message.Cc) & " BCC:" & JoinMailNames(message.Bcc))

            'If isMailToSystem(exToSMTP(JoinMailNames(message.To))) Then

            '    If users.isRequestor(JoinMailboxes(message.From)) Then

            '        'IF MAIL ITEM ID IS NOT IN THE DB
            '        If Not actions.requestMailProcessed(message.MessageId) Then

            '            log("Creating Request: " & message.Subject)

            '            'dictMail CONVERTS THE MAIL INTO DICTIONARY OF STRINGS
            '            'CREATE THE REQUEST INTO THE DATABASE
            '            actions.createNewRequest(dictMail(message))

            '        End If

            '    End If

            'End If

            'Rewrite business rule to create the request

            Dim addressesFrom As String = JoinMailboxes(message.From)
            Dim addresses As String = addressesFrom
            addresses = addresses & ";" & JoinMailboxes(message.Cc)

            Dim addressesList = addresses.Split(New Char() {";"c})

            If isMailToSystem(exToSMTP(JoinMailNames(message.To))) Then
                'Validate if address is a REQUESTOR
                Dim address As String
                For Each address In addressesList
                    If address.Length > 0 Then
                        If users.isRequestor(addressesFrom) Then
                            'IF MAIL ITEM ID IS NOT IN THE DB
                            If Not actions.requestMailProcessed(message.MessageId) Then
                                log("Creating Request: " & message.Subject)

                                'dictMail CONVERTS THE MAIL INTO DICTIONARY OF STRINGS
                                'CREATE THE REQUEST INTO THE DATABASE
                                actions.createNewRequest(dictMail(message))

                            End If
                            'break current loop
                            Exit For
                        End If
                    End If
                Next
            End If

            Dim messageDate As Date = message.Date.Date
            log("message Date: " & messageDate.ToString("MM/dd/yyyy HH:mm"))
            log("last checked email Date: " & dbDate.ToString("MM/dd/yyyy HH:mm"))

            If dbDate.Date < messageDate Then
                dbDate = messageDate
                syscfg.setMailLastCheck(messageDate)
            End If

        Next

        If (imapClient.IsConnected) Then
            imapClient.Disconnect(True)
        End If

        log("CheckMail done")

    End Sub

    Private Function PopulateBody(ByRef reader As StreamReader, ByRef mailData As Dictionary(Of String, String)) As String
        Dim sysconfig As New SysConfig
        Dim urlHost As String = String.Empty
        Dim body As String = String.Empty
        body = reader.ReadToEnd
        For Each kvp As KeyValuePair(Of String, String) In mailData
            If kvp.Key.Substring(0, 1) = "{" Then
                body = body.Replace(kvp.Key, kvp.Value)
            End If
        Next kvp

        'Replace image headers
        urlHost = sysconfig.getSystemUrl()
        body = body.Replace("{sap_logo}", "https://k60.kn3.net/C/C/1/9/3/1/B3B.png")
        '"data:image/png;base64,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'")
        '"cid:logo.png"
        'urlHost + "images/email-templates/logo.png")

        body = body.Replace("{sap_header}", "https://k60.kn3.net/F/2/9/F/F/5/17A.jpg")
        '"data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/4QBARXhpZgAATU0AKgAAAAgAAgESAAMAAAABAAEAAIKYAAIAAAASAAAAJgAAAABQZWRybyBEaWF6IE1vbGlucwD/7AARRHVja3kAAQAEAAAAAwAA/+EFX2h0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8APD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4NCjx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuNi1jMTM4IDc5LjE1OTgyNCwgMjAxNi8wOS8xNC0wMTowOTowMSAgICAgICAgIj4NCgk8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPg0KCQk8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIiB4bWxuczpwaG90b3Nob3A9Imh0dHA6Ly9ucy5hZG9iZS5jb20vcGhvdG9zaG9wLzEuMC8iIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0iQ0U4NjEzQjlBQzU5QjVEOTEwM0NERjM3QjE2NDEyMDgiIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6N0E3RTE3QTBERTc1MTFFNjk0OEJDODJEOUU5NkM5QjciIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6N0E3RTE3OUZERTc1MTFFNjk0OEJDODJEOUU5NkM5QjciIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENTNSBNYWNpbnRvc2giIHBob3Rvc2hvcDpBdXRob3JzUG9zaXRp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")
        '"cid:header.jpg"
        'urlHost + "images/email-templates/header.jpg")

        'Return replaced body
        Return body
    End Function

    Private Sub SendHtmlFormattedEmail(ByVal recipientEmail As String, ByVal subject As String, ByVal body As String, ByVal ai_Id As String)
        Dim dbconn As OleDbConnection
        Dim dbcomm As OleDbCommand
        Dim sql As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New Logging

        dbconn = New OleDbConnection(systemConfig.getConnection)
        dbconn.Open()

        If (String.IsNullOrEmpty(ai_Id)) Then
            sql = "INSERT INTO send (recipients, subject, body) VALUES ('" + recipientEmail + "', '" + subject.Replace("&#34;", """").Replace("'", "&#39;") + "', '" + body.Replace("&#34;", """").Replace("'", "&#39;") + "')"
        Else
            sql = "INSERT INTO send (recipients, subject, body, ai_id) VALUES ('" + recipientEmail + "', '" + subject.Replace("&#34;", """").Replace("'", "&#39;") + "', '" + body.Replace("&#34;", """").Replace("'", "&#39;") + "', " + ai_Id + ")"
        End If

        'sql = "INSERT INTO send (recipients, subject, body, ai_id) VALUES ('" + recipientEmail + "', '" + subject.Replace("&#34;", """").Replace("'", "&#39;") + "', '" + body.Replace("&#34;", """").Replace("'", "&#39;") + "', " + ai_Id + ")"
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
        Dim mailRecepient As String = String.Empty
        Dim ai_Id As String = ""

        mailRecepient = mailData("to")
        If mailData.ContainsKey("{ai_id}") AndAlso Not String.IsNullOrEmpty(mailData("{ai_id}")) Then
            ai_Id = mailData("{ai_id}")
        End If

        mailTemplate = utils.getEmailTemplate(mailData("mail"))

        'GET EMAIL SUBJECT
        'When the subject is provided in the map
        If mailData.ContainsKey("{subject}") AndAlso Not String.IsNullOrEmpty(mailData("{subject}")) Then
            mailSubject = mailData("{subject}")
        Else
            'Get default email subject
            mailSubject = utils.getDefaultEmailSubject(mailData("mail"))
        End If

        Dim reader As StreamReader = New StreamReader(mailTemplate)

        mailBody = PopulateBody(reader, mailData)

        SendHtmlFormattedEmail(mailRecepient, mailSubject, mailBody, ai_Id)
    End Sub

    Sub sendAndDelete()

        Dim waitaMinute As Integer = 0

        Dim Smtp_Server As New SmtpClient
        Dim e_mail As New MailMessage()

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbdel, dbupd, dblst As OleDbCommand
        Dim dbread, dbreadsub As OleDbDataReader
        Dim sql, sqlupd, sqldel, sqlast As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New Logging

        Dim mailSubject As String
        Dim mailBody As String
        Dim checked As Date
        Dim delta As Integer

        Dim ndxStart, ndxEnd As Integer
        Dim aiIdStr As String

        dbconn = New OleDbConnection(systemConfig.getConnection())
        dbconn.Open()

        'CREATE CONNECTION TO SMTP SERVER
        Smtp_Server.UseDefaultCredentials = True
        Smtp_Server.Credentials = New System.Net.NetworkCredential(emailUser, emailPass)
        Smtp_Server.EnableSsl = False
        Smtp_Server.Host = hostSMTP

        'READ INFORMATION FROM DATABASE     
        sql = "SELECT * FROM send;"

        dbcomm = New OleDbCommand(sql, dbconn)
        dbread = dbcomm.ExecuteReader()

        Try

            If dbread.HasRows Then

                While dbread.Read()

                    e_mail = New MailMessage()
                    mailSubject = dbread.GetString(2)
                    mailBody = dbread.GetString(3)

                    log("sendAndDelete - Processing email with id#: " & dbread.GetInt64(0).ToString)

                    'WHEN IS A NEW AI THEN MUST WAIT
                    If mailSubject = "New AI" Then

                        'CHECK IF IT HAS AI ID
                        If dbread.IsDBNull(5) Then

                            'DOES NOT HAVE ID
                            'FIND AI ID IN MAIL BODY LOOK FOR ;;;
                            'WRITE AI ID IN TABLE SEND
                            ndxStart = mailBody.IndexOf(";;;") + 5
                            ndxEnd = mailBody.IndexOf("<", ndxStart)
                            aiIdStr = mailBody.Substring(ndxStart, ndxEnd - ndxStart)

                            If dbread.IsDBNull(4) Then
                                aiIdStr = aiIdStr + ", checked = '" + Now.ToString("yyyy-MM-dd hh:mm:ss tt") + "'"
                            End If

                            sqlupd = "UPDATE send SET ai_id = " + aiIdStr + " WHERE id = " + dbread.GetInt64(0).ToString + ";"
                            dbupd = New OleDbCommand(sqlupd, dbconn)
                            dbupd.ExecuteNonQuery()

                            'MOVE NEXT

                        Else

                            'IS THE ONLY ONE SO CHECK FOR TIMESTAMP GREATER THAN xx MINUTES
                            checked = dbread.GetDateTime(4)
                            delta = DateDiff(DateInterval.Minute, checked, Now)
                            If Math.Abs(delta) >= waitaMinute Then

                                'HAS ID
                                'IF FIND SOME WITH ID GREATER THEN
                                sqlast = "SELECT * FROM send WHERE subject='New AI' AND recipients='" + dbread.GetString(1) + "' AND ai_id = " + dbread.GetInt64(5).ToString + " AND id > " + dbread.GetInt64(0).ToString + ";"
                                dblst = New OleDbCommand(sqlast, dbconn)
                                dbreadsub = dblst.ExecuteReader()

                                If dbreadsub.HasRows Then

                                    'THERE ARE NEW UPDATES SO DELETE THIS ONE
                                    sqldel = "DELETE FROM send WHERE id=" + dbread.GetInt64(0).ToString + ";"
                                    dbdel = New OleDbCommand(sqldel, dbconn)
                                    dbdel.ExecuteNonQuery()

                                    'MOVE NEXT

                                Else

                                    'WAITING ENOUGH MUST SEND NOW
                                    log("sendAndDelete - Email from: " & emailAddressFrom)
                                    log("sendAndDelete - e_mail.To #1: " & CleanInput(dbread.GetString(1)))

                                    e_mail = New MailMessage()
                                    e_mail.From = New System.Net.Mail.MailAddress(emailAddressFrom)
                                    e_mail.To.Add(CleanInput(dbread.GetString(1)))
                                    e_mail.Subject = dbread.GetString(2)
                                    e_mail.IsBodyHtml = True
                                    e_mail.Body = HtmlDecode(dbread.GetString(3))
                                    'Add header and logo images as attachments
                                    'Dim header As New Attachment(".\email-templates\images\email-templates\header.jpg")
                                    'Dim logo As New Attachment(".\email-templates\images\email-templates\logo.png")

                                    'e_mail.Attachments.Add(logo)
                                    'e_mail.Attachments.Add(header)

                                    'Send email
                                    Smtp_Server.Send(e_mail)

                                    'Delete entry in the database
                                    sqldel = "DELETE FROM send WHERE id=" + dbread.GetInt64(0).ToString + ";"
                                    dbdel = New OleDbCommand(sqldel, dbconn)
                                    dbdel.ExecuteNonQuery()

                                    'MOVE NEXT

                                End If

                                dbreadsub.Close()

                            End If

                        End If

                    Else
                        ' Any other subject
                        log("sendAndDelete - Any other subject - Email from:  " & emailAddressFrom)
                        log("sendAndDelete - e_mail.To #1: " & EmailAddressesToString(dbread.GetString(1)))

                        e_mail.From = New System.Net.Mail.MailAddress(emailAddressFrom)
                        e_mail.To.Add(EmailAddressesToString(dbread.GetString(1)))
                        e_mail.Subject = EmailAddressesToString(dbread.GetString(2))

                        'CleanInput
                        e_mail.IsBodyHtml = True
                        e_mail.Body = HtmlDecode(dbread.GetString(3))

                        'Add header and logo images as attachments
                        'Dim header As New Attachment(".\email-templates\images\email-templates\header.jpg")
                        'Dim logo As New Attachment(".\email-templates\images\email-templates\logo.png")

                        'e_mail.Attachments.Add(logo)
                        'e_mail.Attachments.Add(header)

                        'Send email
                        Smtp_Server.Send(e_mail)

                        'Delete entry in the database
                        sqldel = "DELETE FROM send WHERE id=" + dbread.GetInt64(0).ToString + ";"
                        dbdel = New OleDbCommand(sqldel, dbconn)
                        dbdel.ExecuteNonQuery()

                    End If

                End While

            End If

        Catch ex As Exception
            log("Error processing email: " + ex.Message + " - " + ex.StackTrace)
        End Try

        dbread.Close()
        dbconn.Close()

    End Sub

    Private Function CertificateValidationCallBack(ByVal sender As Object, ByVal certificate As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As X509Chain, ByVal policyErrors As SslPolicyErrors) As Boolean
        Return True
    End Function

    Sub closeSession(ByRef imapClient As Net.Imap.ImapClient)
        If (imapClient.IsConnected) Then
            imapClient.Disconnect(True)
        End If
    End Sub

    Sub SendNotifications()
        Dim dbconn As OleDbConnection
        Dim dbcomm, dbupdate As OleDbCommand
        Dim dbread As OleDbDataReader
        Dim sql As String

        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New Logging
        Dim actions As New SapActions

        Dim newMail As New MailTemplate
        Dim mail_dict As New Dictionary(Of String, String)
        Dim log_dict As New Dictionary(Of String, String)

        dbconn = New OleDbConnection(systemConfig.getConnection())
        dbconn.Open()

        'DUE DATE NOTIFICATIONS
        sql = "SELECT * FROM actionitems WHERE due='" + Now.Year.ToString + "-" + Now.Month.ToString + "-" + Now.Day.ToString + "' AND sent_delivery=0;"
        dbcomm = New OleDbCommand(sql, dbconn)

        dbread = dbcomm.ExecuteReader()
        If dbread.HasRows Then
            While dbread.Read()

                'SEND MAIL TO OWNER WITH AIS DETAILS
                mail_dict.Add("mail", "DL") 'AI CREATED
                mail_dict.Add("to", users.getMailById(dbread.GetString(6)))
                mail_dict.Add("{ai_id}", dbread.GetInt64(0).ToString)
                mail_dict.Add("{ai_owner}", users.getNameById(dbread.GetString(6)))
                mail_dict.Add("{description}", dbread.GetString(2)) 'MAIL SUBJECT / AI DESCRIPTION
                mail_dict.Add("{duedate}", utils.formatDateToSTring(Now.Date) & " (Today)")
                mail_dict.Add("{delivery_link}", syscfg.getSystemUrl + "sap_dlvr.aspx?id=" + eLink.enLink(dbread.GetInt64(0).ToString))
                mail_dict.Add("{requestor_name}", users.getNameById(dbread.GetString(6)))
                mail_dict.Add("{app_link}", syscfg.getSystemUrl)
                mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
                mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + dbread.GetInt64(0).ToString)
                mail_dict.Add("{subject}", "Your AI#" & dbread.GetInt64(0).ToString & " is Due Today")

                newMail.SendNotificationMail(mail_dict)

                log_dict.Add("request_id", dbread.GetInt64(1).ToString)
                log_dict.Add("ai_id", dbread.GetInt64(0).ToString)
                log_dict.Add("requestor_id", actions.getRequestorIdFromRequestId(dbread.GetInt64(1)))
                log_dict.Add("admin_id", "system")
                log_dict.Add("owner_id", dbread.GetString(6))
                log_dict.Add("event", "AI_SENT_DELIVERY")
                log_dict.Add("detail", "Some detail here...")

                newLog.LogWrite(log_dict)

                'MARK AS SENT
                sql = "UPDATE actionitems SET sent_delivery=1 WHERE id=" & dbread.GetInt64(0).ToString
                dbupdate = New OleDbCommand(sql, dbconn)
                dbupdate.ExecuteNonQuery()

                mail_dict.Clear()
                log_dict.Clear()

            End While
        End If
        dbread.Close()

        'CONFIRM DUE NOTIFICATIONS
        Dim dueDays As Integer
        Dim ai_id As Long
        Dim twoDays As New TimeSpan(2, 0, 0, 0)
        Dim dueDate As Date = Now.Date.Add(twoDays)
        Dim dayText As String = "day"

        sql = "SELECT * FROM actionitems WHERE due='" + dueDate.Year.ToString + "-" + dueDate.Month.ToString + "-" + dueDate.Day.ToString + "' AND sent_confirm=0;"
        dbcomm = New OleDbCommand(sql, dbconn)

        dbread = dbcomm.ExecuteReader()
        If dbread.HasRows Then
            While dbread.Read()

                ai_id = dbread.GetInt64(0)

                'CALCULATE DUE DAYS
                dueDays = DateDiff(DateInterval.Day, Today.Date, dueDate.Date)
                If dueDays > 1 Then
                    dayText = "days"
                End If

                'SEND MAIL TO OWNER WITH AIS DETAILS
                mail_dict.Add("mail", "CF") 'AI CREATED
                mail_dict.Add("to", users.getMailById(dbread.GetString(6)))
                mail_dict.Add("{ai_id}", ai_id.ToString)
                mail_dict.Add("{description}", dbread.GetString(2)) 'MAIL SUBJECT / AI DESCRIPTION
                mail_dict.Add("{duedate}", dueDate.ToString("dd/MMM/yyyy"))
                mail_dict.Add("{confirm_link}", syscfg.getSystemUrl + "sap_confirm.aspx?id=" + eLink.enLink(ai_id.ToString))
                mail_dict.Add("{extension_link}", syscfg.getSystemUrl + "sap_ext.aspx?id=" + eLink.enLink(ai_id.ToString))
                mail_dict.Add("{need_information}", syscfg.getSystemUrl + "sap_ai_data.aspx?id=" + eLink.enLink(ai_id.ToString))
                mail_dict.Add("{ai_owner}", users.getNameById(dbread.GetString(6)))
                mail_dict.Add("{app_link}", syscfg.getSystemUrl)
                mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
                mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + ai_id.ToString)
                mail_dict.Add("{subject}", "AI#" + ai_id.ToString + " Is due in " + dueDays.ToString + " " + dayText)

                newMail.SendNotificationMail(mail_dict)

                log_dict.Add("request_id", dbread.GetInt64(1).ToString)
                log_dict.Add("ai_id", dbread.GetInt64(0).ToString)
                log_dict.Add("requestor_id", actions.getRequestorIdFromRequestId(dbread.GetInt64(1)))
                log_dict.Add("admin_id", "system")
                log_dict.Add("owner_id", dbread.GetString(6))
                log_dict.Add("event", "AI_SENT_CONFIRM")
                log_dict.Add("detail", "Some detail here...")

                newLog.LogWrite(log_dict)

                'MARK AS SENT
                sql = "UPDATE actionitems SET sent_confirm=1 WHERE id=" & dbread.GetInt64(0).ToString
                dbupdate = New OleDbCommand(sql, dbconn)
                dbupdate.ExecuteNonQuery()

                mail_dict.Clear()
                log_dict.Clear()

            End While
        End If
        dbread.Close()

        dbconn.Close()

    End Sub

    Private Function JoinMailboxes(ByVal mailboxes As MimeKit.InternetAddressList) As String
        Dim builder As New StringBuilder
        Dim index As New Integer

        'Set index to 1
        index = 1

        For Each address As MimeKit.InternetAddress In mailboxes
            If (TypeOf address Is MimeKit.GroupAddress) Then
                Dim group As MimeKit.GroupAddress = CType(address, MimeKit.GroupAddress)
                builder.AppendFormat("{0}: {1};, ", group.Name, JoinMailboxes(group.Members))
            End If
            If (TypeOf address Is MimeKit.MailboxAddress) Then
                Dim email As MimeKit.MailboxAddress = CType(address, MimeKit.MailboxAddress)
                builder.AppendFormat("{0}", email.Address)

                If (index < mailboxes.Count) Then
                    builder.Append(";")
                End If
            End If
            index = index + 1
        Next
        Return builder.ToString()

    End Function

    Private Function JoinMailNames(ByVal addresses As MimeKit.InternetAddressList) As String
        Dim builder As New StringBuilder
        Dim index As New Integer

        'Set index to 1
        index = 1

        For Each address As MimeKit.InternetAddress In addresses
            If (TypeOf address Is MimeKit.GroupAddress) Then
                Dim group As MimeKit.GroupAddress = CType(address, MimeKit.GroupAddress)
                builder.AppendFormat("{0}: {1};, ", group.Name, JoinMailNames(group.Members))
            End If
            If (TypeOf address Is MimeKit.MailboxAddress) Then
                Dim email As MimeKit.MailboxAddress = CType(address, MimeKit.MailboxAddress)
                'builder.AppendFormat("{0} <{1}>, ", email.Name, email.Address)
                builder.AppendFormat("{0}", email.Name)

                If (index < addresses.Count) Then
                    builder.Append(";")
                End If
            End If
            index = index + 1
        Next
        Return builder.ToString()
    End Function

    Private Function GetMailNames(ByVal addresses As MimeKit.InternetAddressList) As List(Of String)
        Dim mailNamesList As New List(Of String)
        Dim builder As New StringBuilder

        For Each address As MimeKit.InternetAddress In addresses
            If (TypeOf address Is MimeKit.GroupAddress) Then
                Dim group As MimeKit.GroupAddress = CType(address, MimeKit.GroupAddress)
                builder.AppendFormat("{0}: {1};, ", group.Name, JoinMailNames(group.Members))
                mailNamesList.Add(builder.ToString)
                builder.Clear()
            End If
            If (TypeOf address Is MimeKit.MailboxAddress) Then
                Dim email As MimeKit.MailboxAddress = CType(address, MimeKit.MailboxAddress)
                mailNamesList.Add(email.Name)
            End If
        Next
        Return mailNamesList
    End Function

    Function CleanInput(strIn As String) As String
        ' Replace invalid characters with empty strings.
        Try
            Return Regex.Replace(strIn, "[^\w\.@-]", "")
            ' If we timeout when replacing invalid characters, 
            ' we should return String.Empty.
        Catch e As Exception
            Return String.Empty
        End Try
    End Function

    Function EmailAddressesToString(emailAddress As String) As String
        Try
            Dim arEmails As New ArrayList
            ' Split string based on comma
            Dim addresses As String() = emailAddress.Split(New Char() {";"c})

            ' Use For Each loop over words and display them

            Dim address As String
            For Each address In addresses
                If address.Length > 0 Then
                    arEmails.Add(address)
                End If
            Next

            Dim strEmails As String = String.Join(", ", arEmails.ToArray)
            Return strEmails
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

End Class
