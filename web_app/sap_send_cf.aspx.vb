Imports System.Data.OleDb
Imports System
Imports System.Collections.Generic
Imports Linker
Imports LogSAPTareas
Imports MailTemplate
Imports SysConfig

Partial Class sap_send_cf
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'CHECK IF DB EXISTS
        '#####TODO###################
        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_ais As OleDbCommand
        Dim dbread_ais As OleDbDataReader
        Dim sql, sql_ais As String

        Dim su As New SapUser
        Dim syscfg As New SysConfig
        Dim link As New Linker

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        Dim ai_id As String = Request("id")

        'CHECK AI ACTUAL STATUS
        sql_ais = "SELECT * FROM actionitems WHERE id=" + ai_id
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)
        dbread_ais = dbcomm_ais.ExecuteReader()

        If dbread_ais.HasRows Then

            dbread_ais.Read()
            Dim request_id As Integer = dbread_ais.GetInt64(1)
            Dim ai_owner As String = dbread_ais.GetString(6)
            Dim ai_descr As String = dbread_ais.GetString(2)
            Dim ai_old_due As Date
            If Not dbread_ais.IsDBNull(4) Then
                ai_old_due = dbread_ais.GetDateTime(4)
            Else
                ai_old_due = Today.Date
            End If
            Dim ai_old_status As String = dbread_ais.GetString(5)
            Dim ai_missing_days As Integer = DateDiff(DateInterval.Day, Today.Date, ai_old_due.Date)

            '////////////////////////////////////////////////////////////
            'CREATE EMAIL AND SEND TO ADMIN
            '////////////////////////////////////////////////////////////
            Dim newMail As New MailTemplate

            Dim mail_dict As New Dictionary(Of String, String)
            mail_dict.Add("mail", "CF") 'AI EXTENSION APPROVED
            mail_dict.Add("to", "ezequielrosa@gmail.com")
            mail_dict.Add("{ai_id}", ai_id)
            mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + ai_id.ToString)
            mail_dict.Add("{description}", ai_descr) 'MAIL SUBJECT / AI DESCRIPTION
            mail_dict.Add("{duedate}", ai_old_due.ToString)
            mail_dict.Add("{confirm_link}", syscfg.getSystemUrl + "sap_ext.aspx?id=" + ai_old_due.ToString)
            mail_dict.Add("{extension_link}", syscfg.getSystemUrl + "sap_ext.aspx?id=" + ai_id.ToString)
            'mail_dict.Add("{requestor_name}", su.getNameById(owner_id))
            mail_dict.Add("{app_link}", syscfg.getSystemUrl)
            mail_dict.Add("{contact_mail_link}", "mailto:" & su.getAdminMail & "?subject=Questions about the report")

            newMail.SendNotificationMail(mail_dict)

            '////////////////////////////////////////////////////////////
            'INSERT LOG HERE
            '////////////////////////////////////////////////////////////

            'EVENT: AI_EXTENSION [R5]

            Dim newLog As New LogSAPTareas

            Dim log_dict As New Dictionary(Of String, String)
            log_dict.Add("ai_id", ai_id)
            log_dict.Add("request_id", request_id.ToString)
            log_dict.Add("owner_id", ai_owner)
            log_dict.Add("event", "AI_SENT_CONFIRM")
            log_dict.Add("detail", "Sent confirmation AI#xx to OWNER...")

            newLog.LogWrite(log_dict)

        End If

        dbconn.Close()

    End Sub
End Class
