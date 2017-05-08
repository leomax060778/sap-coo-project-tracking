Imports System.Data.OleDb
Imports System
Imports System.Collections.Generic
Imports Linker
Imports LogSAPTareas
Imports MailTemplate

Partial Class sap_confirm

    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'CHECK IF DB EXISTS
        '#####TODO###################
        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_ais As OleDbCommand
        Dim dbread_ais As OleDbDataReader
        Dim sql, sql_ais As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'REQUEST ID
        'NEED DECRYPTION HERE
        Dim link As New Linker
        Dim http_ai_id As String = Request("id")
        Dim ai_id As Integer = link.deLink(http_ai_id)
        Dim i As Integer

        If String.IsNullOrEmpty(http_ai_id) Or Not Integer.TryParse(http_ai_id, i) Then
            Response.Redirect(".\sap_error.aspx", False)
        Else
            Integer.TryParse(http_ai_id, ai_id)
        End If

        Dim linked As New Linker

        If i > 1000000 Then
            ai_id = linked.deLink(i.ToString)
        End If

        'CHECK IF AI EXISTS AND AI ACTUAL STATUS IS IN_PROGRESS AND MISSING 2 DAYS
        sql_ais = "SELECT * FROM actionitems WHERE id=" + ai_id.ToString + "AND status='IP' " '+ "AND due=" + Today().Date + 2 dias
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)
        dbread_ais = dbcomm_ais.ExecuteReader()

        If dbread_ais.HasRows Then

            dbread_ais.Read()
            'Dim ai_descr As String = dbread_ais.GetString(2)
            Dim ai_owner As String = dbread_ais.GetString(6)
            'Dim ai_old_due As Date = dbread_ais.GetDateTime(4)
            'Dim ai_extension As Date = dbread_ais.GetDateTime(8)
            'Dim ai_missing_days As Integer = DateDiff(DateInterval.Day, Today.Date, ai_old_due.Date)

            'DUEDATE EXTENSION
            sql = "UPDATE actionitems SET status='CF' WHERE id=" + ai_id.ToString
            dbcomm = New OleDbCommand(sql, dbconn)
            dbcomm.ExecuteScalar()

            '////////////////////////////////////////////////////////////
            'CREATE EMAIL AND SEND TO OWNER
            '////////////////////////////////////////////////////////////
            'Dim newMail As New MailTemplate

            'Dim mail_dict As New Dictionary(Of String, String)
            'mail_dict.Add("mail", "EA") 'AI EXTENSION APPROVED
            'mail_dict.Add("to", "ezequielrosa@gmail.com")
            'mail_dict.Add("{ai_id}", ai_id.ToString)
            'mail_dict.Add("{description}", ai_descr) 'MAIL SUBJECT / AI DESCRIPTION
            'mail_dict.Add("{duedate}", ai_extension.ToString)

            'newMail.SendNotificationMail(mail_dict)

            '////////////////////////////////////////////////////////////
            'INSERT LOG HERE
            '////////////////////////////////////////////////////////////

            'EVENT: AI_EXTENSION [R5]

            Dim newLog As New LogSAPTareas

            Dim log_dict As New Dictionary(Of String, String)
            log_dict.Add("ai_id", ai_id.ToString)
            log_dict.Add("owner_id", ai_owner)
            log_dict.Add("admin_id", users.getId)
            'log_dict.Add("prev_value", ai_old_due.ToString)
            'log_dict.Add("new_value", ai_extension.ToString)
            log_dict.Add("event", "AI_EXTENDED")
            log_dict.Add("detail", "Some detail here...")

            newLog.LogWrite(log_dict)

            '#####TODO###################
            '#####TODO###################
            '#####TODO###################
            Response.Write("Due Date CONFIRMED")
            '#####TODO###################
            '#####TODO###################
            '#####TODO###################

        Else

            'AI DOES NOT EXIST

        End If

        dbread_ais.Close()
        dbconn.Close()

    End Sub

End Class
