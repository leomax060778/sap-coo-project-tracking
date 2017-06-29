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


'Highligth COLOR
'#FFFC6E

Public Class MailTemplate

    Dim Smtp_Server As New SmtpClient
    Dim e_mail As New MailMessage

    'DEFINE GLOBAL VARIABLES
    Public currentEnv As String = "production"
    Public imapServer As String = "imap.global.corp.sap"
    Public imapPort As Integer = 993
    Public isSSL = True
    Public hostSMTP As String = "mail.sap.corp"
    Public emailUser As String = "asa1_sap_mktg_in_ac@global.corp.sap\sap_marketing_in_action"
    Public emailPass As String = "BAoR}:qKQSkzSBO'#4pQ"
    Public emailAddressFrom As String = "sap_marketing_in_action@sap.com"

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

        mailFile = "d:\webapps\test\mails\" & message.MessageId & ".eml"
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
        Dim syscfg As New SysConfig
        Dim result As Boolean = False

        Return mailList.IndexOf(syscfg.getSystemMail) >= 0

    End Function

    Public Sub CheckMail()
        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim actions As New SapActions
        Dim dbDate As Date
        Dim log_dict As New Dictionary(Of String, String)
        Dim newLog As New LogSAPTareas
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
                        log("Marked email for deletion with EntryID: " & email.MessageId & email.Subject)
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
        body = body.Replace("{sap_logo}", urlHost + "images/email-templates/logo.png")
        body = body.Replace("{sap_header}", urlHost + "images/email-templates/header.jpg")

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
        Dim newLog As New LogSAPTareas

        dbconn = New OleDbConnection(syscfg.getConnection)
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
        Dim mailRecepient As String = mailData("to")
        Dim ai_Id As String = ""

        If mailData.ContainsKey("{ai_id}") AndAlso Not String.IsNullOrEmpty(mailData("{ai_id}")) Then
            ai_Id = mailData("{ai_id}")
        End If

        'SELECT MAIL TEMPLATE
        Select Case mailData("mail")
            Case "ND"
                mailTemplate = ".\email-templates\SAP Email A - More info.html"
                mailSubject = "Your request is pending for information lack"
            Case "CF"
                mailTemplate = ".\email-templates\SAP Email B - Due 2 days.html"
                mailSubject = "AI due date confirmation"
            Case "CR"
                mailTemplate = ".\email-templates\SAP Email C - AI Owner.html"
                mailSubject = "New AI"
            Case "NE"
                mailTemplate = ".\email-templates\SAP Email E - Extension Requested.html"
                mailSubject = "AI extension requested"
            Case "EA"
                mailTemplate = ".\email-templates\SAP Email F - Extension Approved.html"
                mailSubject = "AI extension approved"
            Case "ER"
                mailTemplate = ".\email-templates\SAP Email G - Extension Rejected.html"
                mailSubject = "AI extension rejected"
            Case "DL"
                mailTemplate = ".\email-templates\SAP Email I - Due today.html"
                mailSubject = "AI delivery day"
            Case "NR"
                mailTemplate = ".\email-templates\SAP Email H - Admin New Request.html"
                mailSubject = "New RQ created"
            Case "OR"
                mailTemplate = ".\email-templates\SAP Email N - Owner Report.html"
                mailSubject = "Your WIP Action Items"
            Case "AR"
                mailTemplate = ".\email-templates\SAP Email O - Admin Report.html"
                mailSubject = "COO Project Tracking - Admin Report"
            Case Else
                'ERROR / HALT DO NOTHING
        End Select

        Dim reader As StreamReader = New StreamReader(mailTemplate)

        mailBody = PopulateBody(reader, mailData)

        SendHtmlFormattedEmail(mailRecepient, mailSubject, mailBody, ai_Id)

    End Sub

    Sub sendAndDelete()

        Dim waitaMinute As Integer = 1

        Dim Smtp_Server As New SmtpClient
        Dim e_mail As New MailMessage()

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbdel, dbupd, dblst As OleDbCommand
        Dim dbread, dbreadsub As OleDbDataReader
        Dim sql, sqlupd, sqldel, sqlast As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New LogSAPTareas

        Dim mailSubject As String
        Dim mailBody As String
        Dim checked As Date
        Dim delta As Integer

        Dim ndxStart, ndxEnd As Integer
        Dim aiIdStr As String

        dbconn = New OleDbConnection(syscfg.getConnection())
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

                    mailSubject = dbread.GetString(2)
                    mailBody = dbread.GetString(3)

                    log("sendAndDelete - Sending email with subject: " & mailSubject)

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
                                    log("sendAndDelete - Email to: " & CleanInput(dbread.GetString(1)))

                                    e_mail = New MailMessage()
                                    e_mail.From = New System.Net.Mail.MailAddress(emailAddressFrom)
                                    e_mail.To.Add(CleanInput(dbread.GetString(1)))
                                    e_mail.Subject = dbread.GetString(2)
                                    e_mail.IsBodyHtml = True
                                    e_mail.Body = HtmlDecode(dbread.GetString(3))
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
                        log("sendAndDelete - Any other subject - Email from: " & emailAddressFrom)
                        log("sendAndDelete - Any other subject - Email to: " & CleanInput(dbread.GetString(1)))

                        e_mail.From = New System.Net.Mail.MailAddress(emailAddressFrom)
                        e_mail.To.Add(EmailAddressesToString(dbread.GetString(1)))
                        e_mail.Subject = EmailAddressesToString(dbread.GetString(2))

                        'CleanInput
                        e_mail.IsBodyHtml = True
                        e_mail.Body = HtmlDecode(dbread.GetString(3))
                        Smtp_Server.Send(e_mail)

                        'Delete entry in the database
                        sqldel = "DELETE FROM send WHERE id=" + dbread.GetInt64(0).ToString + ";"
                        dbdel = New OleDbCommand(sqldel, dbconn)
                        dbdel.ExecuteNonQuery()

                    End If

                End While

            End If

        Catch ex As Exception
            log("Error when processing email - Unhandled Exception: " + ex.Message)
            log("Error when processing email - Unhandled Exception: " + ex.StackTrace)
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

        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New LogSAPTareas
        Dim actions As New SapActions

        Dim newMail As New MailTemplate
        Dim mail_dict As New Dictionary(Of String, String)
        Dim log_dict As New Dictionary(Of String, String)

        dbconn = New OleDbConnection(syscfg.getConnection())
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
                mail_dict.Add("{duedate}", Now.Date.ToString("dd/MMM/yyyy") & " (Today)")
                mail_dict.Add("{delivery_link}", syscfg.getSystemUrl + "sap_dlvr.aspx?id=" + eLink.enLink(dbread.GetInt64(0).ToString))
                mail_dict.Add("{requestor_name}", users.getNameById(dbread.GetString(6)))
                mail_dict.Add("{app_link}", syscfg.getSystemUrl)
                mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
                mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + dbread.GetInt64(0).ToString)

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
        Dim twoDays As New TimeSpan(2, 0, 0, 0)
        Dim dueDate As Date = Now.Date.Add(twoDays)

        sql = "SELECT * FROM actionitems WHERE due='" + dueDate.Year.ToString + "-" + dueDate.Month.ToString + "-" + dueDate.Day.ToString + "' AND sent_confirm=0;"
        dbcomm = New OleDbCommand(sql, dbconn)

        dbread = dbcomm.ExecuteReader()
        If dbread.HasRows Then
            While dbread.Read()

                'SEND MAIL TO OWNER WITH AIS DETAILS
                mail_dict.Add("mail", "CF") 'AI CREATED
                mail_dict.Add("to", users.getMailById(dbread.GetString(6)))
                mail_dict.Add("{ai_id}", dbread.GetInt64(0).ToString)
                mail_dict.Add("{description}", dbread.GetString(2)) 'MAIL SUBJECT / AI DESCRIPTION
                mail_dict.Add("{duedate}", dueDate.ToString("dd/MMM/yyyy"))
                mail_dict.Add("{confirm_link}", syscfg.getSystemUrl + "sap_confirm.aspx?id=" + eLink.enLink(dbread.GetInt64(0).ToString))
                mail_dict.Add("{extension_link}", syscfg.getSystemUrl + "sap_ext.aspx?id=" + eLink.enLink(dbread.GetInt64(0).ToString))
                mail_dict.Add("{need_information}", syscfg.getSystemUrl + "sap_ai_data.aspx?id=" + eLink.enLink(dbread.GetInt64(0).ToString))
                mail_dict.Add("{ai_owner}", users.getNameById(dbread.GetString(6)))
                mail_dict.Add("{app_link}", syscfg.getSystemUrl)
                mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
                mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + dbread.GetInt64(0).ToString)

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
