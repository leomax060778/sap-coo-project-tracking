Imports System.Data.OleDb
Imports commonLib

Partial Class _Default
    Inherits System.Web.UI.Page

    Dim userCommon As New commonLib.SapUser

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'CHECK IF DB EXISTS
        '#####TODO###################
        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_ais As OleDbCommand
        Dim dbread_ais As OleDbDataReader
        Dim sql, sql_ais As String

        Dim users As New SapUser
        Dim actions As New SapActions
        Dim utils As New Utils
        Dim sysConfiguration As New SystemConfiguration

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        'REQUEST ID
        'NEED DECRYPTION HERE
        Dim link As New Linker
        Dim http_ai_id As String = Request("id")
        Dim ai_id As Integer = link.deLink(http_ai_id)
        Dim i As Integer

        'CHECK IF REQUEST HAS AN AI ID
        If String.IsNullOrEmpty(http_ai_id) Or Not Integer.TryParse(http_ai_id, i) Or (userCommon.getRole <> "RQ" And userCommon.getRole = "OW") Then
            Response.Redirect(".\sap_error.aspx", False)
        Else
            Integer.TryParse(http_ai_id, ai_id)
        End If

        Dim linked As New Linker

        If i > 1000000 Then
            ai_id = linked.deLink(i.ToString)
        End If

        'CHECK IF AI EXISTS AND AI ACTUAL STATUS IS NEED_EXTENSION
        sql_ais = "SELECT * FROM actionitems WHERE id=" + ai_id.ToString + "AND status='DL'"
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)
        dbread_ais = dbcomm_ais.ExecuteReader()

        If dbread_ais.HasRows Then

            If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
                dbread_ais.Read()

                'DUEDATE EXTENSION
                sql = "UPDATE actionitems SET status='IP' WHERE id=" + ai_id.ToString
                dbcomm = New OleDbCommand(sql, dbconn)
                dbcomm.ExecuteScalar()

                actions.requestSetStatus(dbread_ais.GetInt64(1), "", "IP")

                'Create a Lumira AI
                Dim lumira_ai As New Dictionary(Of String, String)
                Dim lumiraReport As New LumiraReports

                lumira_ai.Add("ai_id", ai_id)
                lumira_ai.Add("in_progress", Date.Now.ToString("yyyy/MM/dd HH:mm:ss"))

                lumiraReport.LogActionItemReport(ai_id, lumira_ai)
                'End create Lumira AI

                '////////////////////////////////////////////////////////////
                'CREATE EMAIL AND SEND TO OWNER
                '////////////////////////////////////////////////////////////
                Dim newMail As New MailTemplate

                Dim owner_id As String = dbread_ais.GetString(6)
                Dim request_id As Integer = dbread_ais.GetInt64(1)
                Dim requestor_id As String = actions.getRequestorIdFromRequestId(request_id)
                Dim description As String = dbread_ais.GetString(2)
                Dim duedate As String = utils.formatDateToSTring(dbread_ais.GetDateTime(4))

                Dim mail_dict As New Dictionary(Of String, String)
                mail_dict.Add("mail", "AU") 'AI EXTENSION REJECTED
                mail_dict.Add("to", userCommon.getMailById(owner_id))
                mail_dict.Add("{ai_id}", ai_id.ToString)
                mail_dict.Add("{ai_owner}", userCommon.getNameById(owner_id) & "(" & owner_id & ")")
                mail_dict.Add("{description}", description) 'MAIL SUBJECT / AI DESCRIPTION
                mail_dict.Add("{duedate}", duedate)
                mail_dict.Add("{reason}", Request.Form("reason"))
                mail_dict.Add("{requestor_name}", userCommon.getNameById(requestor_id))
                mail_dict.Add("{app_link}", sysConfiguration.getSystemUrl)
                mail_dict.Add("{contact_mail_link}", "mailto:" & userCommon.getAdminMail & "?subject=Questions about the report")
                mail_dict.Add("{subject}", "AI#" & ai_id.ToString & " Is Incomplete")

                newMail.SendNotificationMail(mail_dict)

                'Sent message to Admins
                Dim adminUsers() As String
                adminUsers = userCommon.getAdminMail().Split(";")

                For Each admin As String In adminUsers
                    Dim adminName = userCommon.getNameByMail(admin)
                    Dim adminId = userCommon.getIdByMail(admin)

                    mail_dict = New Dictionary(Of String, String)
                    mail_dict.Add("to", admin)
                    mail_dict.Add("{ai_id}", ai_id.ToString)
                    mail_dict.Add("{ai_owner}", adminName & "(" & adminId & ")")
                    mail_dict.Add("{description}", description) 'MAIL SUBJECT / AI DESCRIPTION
                    mail_dict.Add("{duedate}", duedate)
                    mail_dict.Add("{reason}", Request.Form("reason"))
                    mail_dict.Add("{requestor_name}", userCommon.getNameById(requestor_id))
                    mail_dict.Add("{app_link}", sysConfiguration.getSystemUrl)
                    mail_dict.Add("{contact_mail_link}", "mailto:" & userCommon.getAdminMail & "?subject=Questions about the report")
                    mail_dict.Add("{subject}", "AI#" & ai_id.ToString & " Is Incomplete")

                    newMail.SendNotificationMail(mail_dict)

                Next

                '////////////////////////////////////////////////////////////
                'INSERT LOG HERE
                '////////////////////////////////////////////////////////////

                'EVENT: AI_EXTENSION [R5]

                Dim newLog As New Logging

                Dim log_dict As New Dictionary(Of String, String)
                log_dict.Add("ai_id", ai_id.ToString)
                log_dict.Add("request_id", request_id.ToString)
                log_dict.Add("admin_id", userCommon.getId)
                log_dict.Add("owner_id", owner_id)
                log_dict.Add("requestor_id", requestor_id)
                log_dict.Add("event", "AI_UNCOMPLETED")
                log_dict.Add("detail", Request.Form("reason"))

                newLog.LogWrite(log_dict)

                Dim redirectTo As String
                redirectTo = sysConfiguration.getSystemUrl + "sap_main.aspx"
                Response.Redirect(redirectTo, False)
            End If
        Else

            'AI DOES NOT EXIST

        End If

        dbread_ais.Close()
        dbconn.Close()

    End Sub
End Class
