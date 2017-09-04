Imports System.Data.OleDb
Imports common
Imports commonLib

Partial Class sap_accept_due
    Inherits System.Web.UI.Page

    Dim sysConfiguration As New SystemConfiguration
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

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        'REQUEST ID
        'NEED DECRYPTION HERE
        Dim link As New Linker
        Dim http_ai_id As String = Request("id")
        Dim ai_id As Integer '= link.deLink(http_ai_id)
        Dim i As Integer

        'CHECK IF REQUEST HAS AN AI ID
        If String.IsNullOrEmpty(http_ai_id) Or Not Integer.TryParse(http_ai_id, i) Then
            Response.Redirect(".\sap_error.aspx", False)
        Else
            Integer.TryParse(http_ai_id, ai_id)
        End If

        Dim linked As New Linker

        If i > 1000000 Then
            ai_id = linked.deLink(http_ai_id)
        End If

        'CHECK IF AI EXISTS AND AI ACTUAL STATUS IS NEED_EXTENSION
        sql_ais = "SELECT * FROM actionitems WHERE id=" + ai_id.ToString '+ " AND extension IS NOT NULL"
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)
        dbread_ais = dbcomm_ais.ExecuteReader()

        If dbread_ais.HasRows Then

            dbread_ais.Read()
            Dim ai_descr As String = dbread_ais.GetString(2)
            Dim ai_owner As String = dbread_ais.GetString(6)
            Dim ai_old_due As Date
            If Not dbread_ais.IsDBNull(4) Then
                ai_old_due = dbread_ais.GetDateTime(4)
            Else
                ai_old_due = Today.Date
            End If
            Dim ai_extension As Date
            If Not dbread_ais.IsDBNull(8) Then
                ai_extension = dbread_ais.GetDateTime(8)
            Else
                ai_extension = Today.Date
            End If
            Dim request_id As Integer = dbread_ais.GetInt64(1)

            'DUEDATE EXTENSION
            sql = "UPDATE actionitems SET due='" + ai_extension.Year.ToString + "-" + ai_extension.Month.ToString + "-" + ai_extension.Day.ToString + "', extension=NULL, status='IP' WHERE id=" + ai_id.ToString
            dbcomm = New OleDbCommand(sql, dbconn)
            dbcomm.ExecuteScalar()

            '////////////////////////////////////////////////////////////
            'CREATE EMAIL AND SEND TO OWNER
            '////////////////////////////////////////////////////////////
            Dim newMail As New MailTemplate

            Dim mail_dict As New Dictionary(Of String, String)
            mail_dict.Add("mail", "EA") 'AI EXTENSION APPROVED
            mail_dict.Add("to", userCommon.getMailById(ai_owner))
            mail_dict.Add("{ai_id}", ai_id.ToString)
            mail_dict.Add("{description}", ai_descr) 'MAIL SUBJECT / AI DESCRIPTION
            mail_dict.Add("{duedate}", utils.formatDateToSTring(ai_extension))
            mail_dict.Add("{ai_owner}", userCommon.getNameById(ai_owner))
            mail_dict.Add("{app_link}", sysConfiguration.getSystemUrl)
            mail_dict.Add("{contact_mail_link}", "mailto:" & userCommon.getAdminMail & "?subject=Questions about the report")
            mail_dict.Add("{subject}", "Extension for AI#" & ai_id.ToString & " has been approved")

            newMail.SendNotificationMail(mail_dict)

            '////////////////////////////////////////////////////////////
            'INSERT LOG HERE
            '////////////////////////////////////////////////////////////

            'EVENT: AI_EXTENSION [R5]

            Dim newLog As New Logging
            Dim log_dict As New Dictionary(Of String, String)

            log_dict.Add("ai_id", ai_id.ToString)
            log_dict.Add("request_id", request_id.ToString)
            log_dict.Add("admin_id", userCommon.getId)
            log_dict.Add("owner_id", ai_owner)
            log_dict.Add("requestor_id", actions.getRequestorIdFromRequestId(request_id))
            log_dict.Add("prev_value", ai_old_due.ToString)
            log_dict.Add("new_value", ai_extension.ToString)
            log_dict.Add("event", "AI_EXTENDED")
            log_dict.Add("detail", "Extension approved")

            newLog.LogWrite(log_dict)

        Else

            'AI DOES NOT EXIST

        End If

        dbread_ais.Close()
        dbconn.Close()

        Response.Redirect(sysConfiguration.getSystemUrl + "sap_main.aspx", False)
    End Sub

End Class
