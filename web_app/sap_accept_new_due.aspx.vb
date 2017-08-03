Imports System.Data.OleDb
Imports System
Imports System.Collections.Generic
Imports Linker
Imports LogSAPTareas
Imports MailTemplate
Imports common
Imports commonLib

Partial Class sap_accept_new_due
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
        Dim actions As New SapActions

        dbconn = New OleDbConnection(syscfg.getConnection)
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

            'DUEDATE ACCEPT
            sql = "UPDATE actionitems SET status='IP' WHERE id=" + ai_id.ToString
            dbcomm = New OleDbCommand(sql, dbconn)
            dbcomm.ExecuteScalar()

            Dim newLog As New LogSAPTareas
            Dim log_dict As New Dictionary(Of String, String)

            log_dict.Add("ai_id", ai_id.ToString)
            log_dict.Add("request_id", dbread_ais.GetInt64(1).ToString)
            log_dict.Add("admin_id", users.getId)
            log_dict.Add("owner_id", dbread_ais.GetString(6))
            log_dict.Add("requestor_id", actions.getRequestorIdFromRequestId(dbread_ais.GetInt64(1)))
            log_dict.Add("prev_value", "PD")
            log_dict.Add("new_value", "IP")
            log_dict.Add("event", "AI_ACCEPTED")
            log_dict.Add("detail", "New due date accepted")

            newLog.LogWrite(log_dict)

        Else

            'AI DOES NOT EXIST

        End If

        dbread_ais.Close()
        dbconn.Close()

        Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
    End Sub

End Class
