﻿Imports System.Data.OleDb
Imports System
Imports System.Collections.Generic
Imports Linker
Imports LogSAPTareas
Imports MailTemplate

Partial Class sap_reject_due
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
        Dim ai_id As Integer = link.deLink(http_ai_id)
        Dim i As Integer

        'CHECK IF REQUEST HAS AN AI ID
        If String.IsNullOrEmpty(http_ai_id) Or Not Integer.TryParse(http_ai_id, i) Then
            Response.Redirect(".\sap_error.aspx", False)
        Else
            Integer.TryParse(http_ai_id, ai_id)
        End If

        Dim linked As New Linker

        If i > 1000000 Then
            ai_id = linked.deLink(i.ToString)
        End If

        'CHECK IF AI EXISTS AND AI ACTUAL STATUS IS NEED_EXTENSION
        sql_ais = "SELECT * FROM actionitems WHERE id=" + ai_id.ToString + "AND status='NE'"
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)
        dbread_ais = dbcomm_ais.ExecuteReader()

        If dbread_ais.HasRows Then

            If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
                dbread_ais.Read()
                Dim ai_descr As String = dbread_ais.GetString(2)
                Dim ai_owner As String = dbread_ais.GetString(6)
                Dim ai_duedate As Date
                If Not dbread_ais.IsDBNull(4) Then
                    ai_duedate = dbread_ais.GetDateTime(4)
                Else
                    ai_duedate = Today.Date
                End If
                Dim request_id As Integer = dbread_ais.GetInt64(1)
                Dim requestor_id As String = actions.getRequestorIdFromRequestId(request_id)
                'Dim ai_missing_days As Integer = DateDiff(DateInterval.Day, Today.Date, ai_old_due.Date)

                'DUEDATE EXTENSION
                sql = "UPDATE actionitems SET status='IP' WHERE id=" + ai_id.ToString
                dbcomm = New OleDbCommand(sql, dbconn)
                dbcomm.ExecuteScalar()

                '////////////////////////////////////////////////////////////
                'CREATE EMAIL AND SEND TO OWNER
                '////////////////////////////////////////////////////////////
                Dim newMail As New MailTemplate

                Dim mail_dict As New Dictionary(Of String, String)
                mail_dict.Add("mail", "ER") 'AI EXTENSION REJECTED
                mail_dict.Add("to", users.getMailById(ai_owner))
                mail_dict.Add("{ai_id}", ai_id.ToString)
                mail_dict.Add("{ai_owner}", users.getNameById(ai_owner) & "(" & ai_owner & ")")
                mail_dict.Add("{requestor}", users.getMailById(requestor_id))
                mail_dict.Add("{description}", ai_descr) 'MAIL SUBJECT / AI DESCRIPTION
                mail_dict.Add("{duedate}", ai_duedate.ToString("dd-MM-yyyy"))
                mail_dict.Add("{reason}", Request.Form("reason"))
                mail_dict.Add("{requestor_name}", users.getNameById(requestor_id))
                mail_dict.Add("{app_link}", syscfg.getSystemUrl)
                mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")

                newMail.SendNotificationMail(mail_dict)

                '////////////////////////////////////////////////////////////
                'INSERT LOG HERE
                '////////////////////////////////////////////////////////////

                'EVENT: AI_EXTENSION [R5]

                Dim newLog As New LogSAPTareas

                Dim log_dict As New Dictionary(Of String, String)
                log_dict.Add("ai_id", ai_id.ToString)
                log_dict.Add("request_id", request_id.ToString)
                log_dict.Add("admin_id", users.getId)
                log_dict.Add("owner_id", ai_owner)
                log_dict.Add("requestor_id", requestor_id)
                log_dict.Add("event", "AI_NOT_EXTENDED")
                log_dict.Add("detail", Request.Form("reason"))

                newLog.LogWrite(log_dict)

                'Set message on label
                lblMessage.Text = "Succesfully delivered your feedback"

                Dim redirectTo As String
                redirectTo = syscfg.getSystemUrl + "sap_main.aspx"
                Response.Redirect("./sap_main.aspx", False)

            End If
        Else
            'AI DOES NOT EXIST
            'Set message on label
            lblMessage.Text = "Action Item does not exist in the system. Please try again!"

        End If

        dbread_ais.Close()
        dbconn.Close()

    End Sub
End Class
