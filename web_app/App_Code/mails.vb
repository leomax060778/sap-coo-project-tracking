Imports System.IO
Imports System.Data.OleDb

Public Class MailTemplate

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
            result = System.Text.RegularExpressions.Regex.Replace(result, "( )+", " ")

            ' Remove the header (prepare first by clearing attributes)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*head([^>])*>", "<head>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*head( )*>)", "</head>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<head>).*(</head>)", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' remove all scripts (prepare first by clearing attributes)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*script([^>])*>", "<script>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*script( )*>)", "</script>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            'result = System.Text.RegularExpressions.Regex.Replace(result,
            '         @"(<script>)([^(<script>\.</script>)])*(</script>)",
            '         string.Empty,
            '         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<script>).*(</script>)", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' remove all styles (prepare first by clearing attributes)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*style([^>])*>", "<style>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*style( )*>)", "</style>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<style>).*(</style>)", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' insert tabs in spaces of <td> tags
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*td([^>])*>", vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' insert line breaks in places of <BR> and <LI> tags
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*br( )*>", vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*li( )*>", vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' insert line paragraphs (double line breaks) in place
            ' if <P>, <DIV> and <TR> tags
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*div([^>])*>", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*tr([^>])*>", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*p([^>])*>", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' Remove remaining tags like <a>, links, images,
            ' comments etc - anything that's enclosed inside < >
            result = System.Text.RegularExpressions.Regex.Replace(result, "<[^>]*>", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' replace special characters:
            result = System.Text.RegularExpressions.Regex.Replace(result, " ", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            result = System.Text.RegularExpressions.Regex.Replace(result, "&bull;", " * ", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&lsaquo;", "<", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&rsaquo;", ">", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&trade;", "(tm)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&frasl;", "/", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&lt;", "<", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&gt;", ">", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&copy;", "(c)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&reg;", "(r)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            ' Remove all others. More can be added, see
            ' http://hotwired.lycos.com/webmonkey/reference/special_characters/
            result = System.Text.RegularExpressions.Regex.Replace(result, "&(.{2,6});", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

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
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbCr & ")( )+(" & vbCr & ")", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbTab & ")( )+(" & vbTab & ")", vbTab & vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbTab & ")( )+(" & vbCr & ")", vbTab & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbCr & ")( )+(" & vbTab & ")", vbCr & vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            ' Remove redundant tabs
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbCr & ")(" & vbTab & ")+(" & vbCr & ")", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            ' Remove multiple tabs following a line break with just one tab
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbCr & ")(" & vbTab & ")+", vbCr & vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
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


    Public Function CheckMail() As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)
        Return result
    End Function


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
        body = body.Replace("{sap_logo}", urlHost + "images/logo.png")
        body = body.Replace("{sap_header}", urlHost + "images/header.jpg")

        'Return replaced body
        Return body
    End Function


    Private Sub SendHtmlFormattedEmail(ByVal recipientEmail As String, ByVal subject As String, ByVal body As String)

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_ais As OleDbCommand
        Dim dbread_ais As OleDbDataReader
        Dim sql, sql_ais As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New LogSAPTareas

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql = "INSERT INTO send (recipients, subject, body) VALUES ('" + recipientEmail + "', '" + subject.Replace("&#34;", """").Replace("'", "&#39;") + "', '" + body.Replace("&#34;", """").Replace("'", "&#39;") + "')"
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

        'SELECT MAIL TEMPLATE
        Select Case mailData("mail")
            Case "ND"
                mailTemplate = "~/email-templates/SAP Email A - More info.html"
                mailSubject = "Your request is pending for information lack"
            Case "CF"
                mailTemplate = "~/email-templates/SAP Email B - Due 2 days.html"
                mailSubject = "AI due date confirmation"
            Case "CR"
                mailTemplate = "~/email-templates/SAP Email C - AI Owner.html"
                mailSubject = "New AI"
            Case "CP"
                mailTemplate = "~/email-templates/SAP Email D - Delivery Approval.html"
                mailSubject = "Delivery approval"
            Case "NE"
                mailTemplate = "~/email-templates/SAP Email E - Extension Requested.html"
                mailSubject = "AI extension requested"
            Case "EA"
                mailTemplate = "~/email-templates/SAP Email F - Extension Approved.html"
                mailSubject = "AI extension approved"
            Case "ER"
                mailTemplate = "~/email-templates/SAP Email G - Extension Rejected.html"
                mailSubject = "AI extension rejected"
            Case "NR"
                mailTemplate = "~/email-templates/SAP Email H - Admin New Request.html"
                mailSubject = "New RQ created"
            Case "DL"
                mailTemplate = "~/email-templates/SAP Email I - Due today.html"
                mailSubject = "AI delivery day"
            Case "AC"
                mailTemplate = "~/email-templates/SAP Email J - AI Completed.html"
                mailSubject = "AI completed"
            Case "AU"
                mailTemplate = "~/email-templates/SAP Email K - AI Uncompleted.html"
                mailSubject = "AI uncompleted"
            Case Else
                'ERROR / HALT DO NOTHING
        End Select

        Dim reader As StreamReader = New StreamReader(HttpContext.Current.Server.MapPath(mailTemplate))

        mailBody = PopulateBody(reader, mailData)

        SendHtmlFormattedEmail(mailRecepient, mailSubject, mailBody)

    End Sub

End Class