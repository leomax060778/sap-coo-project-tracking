Imports System.Data.OleDb
Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.HttpUtility
Imports System.Collections.Generic
Imports SysConfig
Imports SapActions
'Imports MailTemplate

Partial Class _Default
    Inherits System.Web.UI.Page

    Private Function HumanizeFwd(ByVal i As Integer) As String
        Dim result As String
        Select Case i
            Case Is < 0
                result = "OverDue"
            Case 0
                result = "Today"
            Case 1
                result = "Tomorrow"
            Case Else
                result = "within " & i.ToString & " days"
        End Select
        Return result
    End Function

    Private Function HumanizeBkw(ByVal i As Integer) As String
        Dim result As String
        Select Case i
            Case Is < 0
                result = "Error"
            Case 0
                result = "Today"
            Case 1
                result = "Yesterday"
            Case Else
                result = i.ToString & " days ago"
        End Select
        Return result
    End Function

    Private Function ai_Str_Status(ByVal s As String) As String
        Dim result As String
        Select Case s
            Case "PD"
                result = "Pending"
            Case "IP"
                result = "In Progress"
            Case "NE"
                result = "Extending"
            Case "OD"
                result = "Overdue"
            Case "CF"
                result = "Confirmed"
            Case "DL"
                result = "Delivered"
            Case Else
                result = "Unset"
        End Select
        Return result
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim syscfg As New SysConfig

        Dim su As New SapUser
        Dim ro As String = su.getRole()

        Dim actions As New SapActions

        current_user.Text = su.getName()

        If ro = "OW" Then
            Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
        Else

            Dim dbconn As OleDbConnection
            Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
            Dim dbread_req, dbread_ais As OleDbDataReader
            Dim sql, sql_req, sql_ais As String

            '#####TODO:#CHECK#IF#DB#EXIST###########

            dbconn = New OleDbConnection(syscfg.getConnection)
            dbconn.Open()

            'REQUEST ID
            Dim http_req_id As String = Request("id")
            Dim request_id As Integer
            Dim i As Integer

            If String.IsNullOrEmpty(http_req_id) Or Not Integer.TryParse(http_req_id, i) Then
                Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
            Else
                Integer.TryParse(http_req_id, request_id)
            End If

            'CLEAR FORM
            descr.InnerText = ""
            duedate.Value = ""

            'REQUEST FORM
            If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

                '#####TODO:#CHECK#VALUES################
                Dim http_req_form_subj As String = Request.Form("subj")
                Dim http_req_form_descr As String = Request.Form("descr")
                Dim http_req_form_duedate As String = Request.Form("duedate")
                Dim changes As String = ""
                Dim due_updated As Boolean = False
                Dim info_updated As Boolean = False

                'GET DB ACTUAL DATA
                sql_req = "SELECT * FROM requests WHERE id=" + request_id.ToString
                dbcomm_req = New OleDbCommand(sql_req, dbconn)
                dbread_req = dbcomm_req.ExecuteReader()

                If dbread_req.HasRows Then
                    dbread_req.Read()

                    'IF DB DATA IS DIFFERENT FROM REQUEST THEN
                    If dbread_req.GetString(4) <> http_req_form_subj Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "subject='" + http_req_form_subj.Replace("&#34;", """").Replace("'", "&#39;") + "'"
                        info_updated = True
                    End If

                    If dbread_req.GetString(5) <> http_req_form_descr Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "detail='" + http_req_form_descr.Replace("&#34;", """").Replace("'", "&#39;") + "'"
                        info_updated = True
                    End If

                    Dim duedb As Date
                    If Not dbread_req.IsDBNull(6) Then
                        duedb = dbread_req.GetDateTime(6)
                    Else
                        duedb = Nothing
                    End If

                    Dim duereqdate As Date = Date.ParseExact(http_req_form_duedate.Replace(" ", ""), "dd/MMM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None)
                    'Date("#" & http_req_form_duedate.Replace(" ", "") & "#")
                    'Dim duereq() As String = http_req_form_duedate.Split("-")
                    If duedb.Year <> duereqdate.Year Or duedb.Month <> duereqdate.Month Or duedb.Day <> duereqdate.Day Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "due='" + duereqdate.Year.ToString + "-" + duereqdate.Month.ToString + "-" + duereqdate.Day.ToString + "'"
                        due_updated = True
                    End If

                    'IF THERE ARE CHANGES TO BE DONE
                    If Not String.IsNullOrEmpty(changes) Then
                        sql_req = "UPDATE requests SET " + changes + " WHERE id=" + request_id.ToString
                        dbcomm_req = New OleDbCommand(sql_req, dbconn)
                        dbcomm_req.ExecuteNonQuery()

                        'IF CHANGED DUE DATE THEN REQUEST STATUS CHANGES TO CREATED
                        'AND MUST UPDATE ALL THE AI WITH DUE DATE GREATER THAN UPDATE
                        'AND SEND NOTIFICATION TO ALL THE OWNERS
                        If due_updated Then

                            actions.requestSetStatus(request_id.ToString, "PD", "CR")

                            'GET DB ACTUAL DATA
                            'sql_ais = "UPDATE actionitems SET due='" + duereqdate.Year.ToString + "-" + duereqdate.Month.ToString + "-" + duereqdate.Day.ToString + "' WHERE request_id=" + request_id.ToString + " AND due > '" + duereqdate.Year.ToString + "-" + duereqdate.Month.ToString + "-" + duereqdate.Day.ToString + "'"
                            'dbcomm_ais = New OleDbCommand(sql_ais, dbconn)
                            'dbcomm_ais.ExecuteNonQuery()
                        End If

                        '###########################################################
                        '#TODO NOTIFY EVERYONE
                        '#TODO CREATE LOG
                        '###########################################################

                        'IF THE RQ DUE DATE IS UNSET THEN SET IT TO THE AI DUE DATE
                        'AND CHANGE REQUEST STATUS TO CREATED
                        'If actions.requestIsUnset(http_req_id) Then
                        '    sql_req = "UPDATE requests SET status='CR', due='" + http_req_form_duedate + "' WHERE id=" + request_id.ToString
                        'Else
                        '    sql_req = "UPDATE requests SET status='CR' WHERE id=" + request_id.ToString
                        'End If
                        'dbcomm_req = New OleDbCommand(sql_req, dbconn)
                        'dbcomm_req.ExecuteNonQuery()

                        'NEED ENCRYPTION HERE
                        'Dim link As New Linker

                        '////////////////////////////////////////////////////////////
                        'CREATE EMAIL AND SEND TO OWNER
                        '////////////////////////////////////////////////////////////
                        'Dim newMail As New MailTemplate

                        'Dim mail_dict As New Dictionary(Of String, String)
                        'mail_dict.Add("mail", "CR") 'NEW AI CREATED
                        'mail_dict.Add("to", "ezequielrosa@gmail.com")
                        'mail_dict.Add("{ai_id}", ai_id.ToString)
                        'mail_dict.Add("{description}", http_req_form_descr) 'MAIL SUBJECT / AI DESCRIPTION
                        'mail_dict.Add("{duedate}", http_req_form_duedate)
                        'mail_dict.Add("{accept_link}", syscfg.getSystemUrl + "sap_accept_due.aspx?id=" + link.enLink(ai_id.ToString))
                        'mail_dict.Add("{extension_link}", syscfg.getSystemUrl + "sap_ext.aspx?id=" + link.enLink(ai_id.ToString))

                        'newMail.SendNotificationMail(mail_dict)

                        '////////////////////////////////////////////////////////////
                        'INSERT LOG HERE
                        '////////////////////////////////////////////////////////////
                        'ai_id, request_id, owner_id, requestor_id, admin_id, event, prev_value, new_value, detail

                        'EVENT: AI_CREATED [R1]

                        Dim newLog As New LogSAPTareas

                        changes = ""

                        If due_updated Then
                            changes = "Info changed "
                        End If

                        If info_updated Then
                            changes = changes + "Due Date changed"
                        End If

                        Dim log_dict As New Dictionary(Of String, String)
                        'log_dict.Add("ai_id", ai_id.ToString)
                        log_dict.Add("request_id", request_id.ToString)
                        log_dict.Add("admin_id", su.getId)
                        'log_dict.Add("owner_id", "OWNER_ID")
                        log_dict.Add("event", "RQ_UPDATED")
                        log_dict.Add("detail", changes)

                        newLog.LogWrite(log_dict)

                    End If

                End If

                Dim redirectTo As String
                redirectTo = syscfg.getSystemUrl + "sap_req.aspx?id=" + http_req_id
                Response.Redirect(redirectTo, False)
            End If

            Dim extra As String = ""

            'ROWS ITERATION
            sql_req = "SELECT * FROM requests WHERE id=" + request_id.ToString
            'sql_ais = "SELECT * FROM actionitems WHERE request_id=" + request_id.ToString + extra

            dbcomm_req = New OleDbCommand(sql_req, dbconn)
            'dbcomm_ais = New OleDbCommand(sql_ais, dbconn)

            'ROWS ITERATION
            '#####TODO:#CHECK#IF#REQUEST#EXIST######
            dbread_req = dbcomm_req.ExecuteReader()
            If dbread_req.HasRows Then

                dbread_req.Read()

                req_id.Text = req_id.Text + Convert.ToInt64(dbread_req.GetValue(0)).ToString
                'ainumber.Text = "#" + req_id.Text

                req_detail.Text = dbread_req.GetString(4).Replace("&#39;", "'")
                subj.InnerText = req_detail.Text

                'req_description.Text = dbread_req.GetString(5).Replace("&#39;", "'")

                'If Len(req_description.Text) > 250 Then
                '    req_description.Text = Mid(req_description.Text, 1, 250) & "..."
                'End If

                Dim mailHTML As String = dbread_req.GetString(3)
                mail_detail.InnerHtml = HtmlDecode(mailHTML.Replace("&#34;", """").Replace("&#39;", "'"))

                descr.InnerText = dbread_req.GetString(5).Replace("&#39;", "'")

                Dim req_created As Date
                Dim req_duedate As Date
                req_created = dbread_req.GetDateTime(2)
                If dbread_req.IsDBNull(6) Then
                    req_duedate = Today()
                Else
                    req_duedate = dbread_req.GetDateTime(6)
                End If

                duedate.Value = req_duedate.ToString("dd/MMM/yyyy") 'req_duedate.Year.ToString + "-" + req_duedate.Month.ToString + "-" + req_duedate.Day.ToString

				link_del_req.HRef = syscfg.getSystemUrl + "sap_req_del.aspx?id=" + Convert.ToInt64(dbread_req.GetValue(0)).ToString

            End If

            dbread_req.Close()
            dbconn.Close()

            End If

    End Sub

End Class
