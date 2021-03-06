﻿Imports System.Data.OleDb
Imports System.Web.HttpUtility
Imports commonLib
'Imports MailTemplate

Partial Class _Default
    Inherits System.Web.UI.Page

    Dim utils As New Utils
    Dim sysConfiguration As New SystemConfiguration
    Dim userCommon As New commonLib.SapUser

    Private Sub CommandBtn_Click(ByVal sender As Object, ByVal e As CommandEventArgs)

        Dim actions As New SapActions
        Dim ro As String = userCommon.getRole()

        If ro = "OW" Then
            Response.Redirect(sysConfiguration.getSystemUrl + "sap_main.aspx", False)
        Else
            Dim dbconn As OleDbConnection
            Dim dbcomm_req As OleDbCommand
            Dim sql_req As String
            Dim rq_id As String = CType(e.CommandArgument, String)

            '#####TODO:#CHECK#IF#DB#EXIST###########

            dbconn = New OleDbConnection(sysConfiguration.getConnection)
            dbconn.Open()

            'LOG INFORMATION
            Dim log_rq_id As String = rq_id
            Dim log_owner As String = "Current USER here..."
            Dim log_record As Boolean = False

            If e.CommandName = "X" Then
                'CHANGE AIS STATUS TO CANCELED
                sql_req = "UPDATE actionitems SET status='XX' WHERE id=" + rq_id
                dbcomm_req = New OleDbCommand(sql_req, dbconn)
                dbcomm_req.ExecuteNonQuery()

                '###################################################
                'WOULD HAVE TO SEND EMAIL TO EACH OWNER TO INFORMING
                'WOULD HAVE TO LOG CHANGE STATUS FROM ?? TO XX
                '###################################################

                Dim newLog As New Logging

                Dim log_dict As New Dictionary(Of String, String)
                log_dict.Add("request_id", rq_id)
                log_dict.Add("admin_id", userCommon.getId)
                log_dict.Add("event", "AI_CANCELED")
                log_dict.Add("detail", "Action Item Canceled")

                newLog.LogWrite(log_dict)

                Response.Redirect(sysConfiguration.getSystemUrl + "sap_req.aspx?id=" + Request("id"), False)
            End If

            dbconn.Close()

        End If

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim link As New Linker

        current_user.Text = userCommon.getFullName()

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
        Dim dbread_req, dbread_ais As OleDbDataReader
        Dim sql, sql_req, sql_ais As String

        '#####TODO:#CHECK#IF#DB#EXIST###########
        Dim actions As New SapActions

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        'REQUEST ID
        Dim http_req_id As String = Request("id")
        Dim request_id As Integer
        Dim i As Integer

        If String.IsNullOrEmpty(http_req_id) Or Not Integer.TryParse(http_req_id, i) Then
            Response.Redirect(sysConfiguration.getSystemUrl + "sap_main.aspx", False)
        Else
            Integer.TryParse(http_req_id, request_id)
        End If

        'CLEAR FORM
        descr.InnerText = ""
        duedate.Value = ""
        owner.Value = ""

        'REQUEST FORM
        If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

            '#####TODO:#CHECK#VALUES################
            Dim http_req_form_descr As String = Request.Form("descr")
            Dim http_req_form_owner As String = Request.Form("owner")
            Dim http_req_form_duedate As String = Request.Form("duedate")

            'ID LAST AI CREATED
            Dim ai_id As Long

            'INSERT NEW ROW
            sql = "INSERT INTO actionitems (request_id, description, owner, due, original, status) VALUES (" + http_req_id + ", '" + http_req_form_descr.Replace("'", "&#39;") + "', '" + http_req_form_owner + "', '" + http_req_form_duedate + "', '" + http_req_form_duedate + "', 'PD')"
            dbcomm = New OleDbCommand(sql, dbconn)
            dbcomm.ExecuteNonQuery()
            dbcomm.CommandText = "SELECT @@IDENTITY"
            ai_id = dbcomm.ExecuteScalar()

            'Create a Lumira AI
            Dim lumira_ai As New Dictionary(Of String, String)
            Dim lumiraReport As New LumiraReports

            lumira_ai.Add("ai_id", ai_id)
            lumira_ai.Add("req_id", http_req_id)
            lumira_ai.Add("owner", http_req_form_owner)
            lumira_ai.Add("description", http_req_form_descr.Replace("'", "&#39;"))
            lumira_ai.Add("due", http_req_form_duedate)
            lumira_ai.Add("original_due", http_req_form_duedate)
            lumira_ai.Add("created", Date.Now.ToString("yyyy/MM/dd HH:mm:ss"))

            lumiraReport.LogActionItemReport(ai_id, lumira_ai)
            'End create Lumira AI

            'IF THE RQ DUE DATE IS UNSET THEN SET IT TO THE AI DUE DATE
            'AND CHANGE REQUEST STATUS TO CREATED
            Dim lumira_request As New Dictionary(Of String, String)

            If actions.requestIsUnset(http_req_id) Then
                sql_req = "UPDATE requests SET status='CR', due='" + http_req_form_duedate + "' WHERE id=" + request_id.ToString

                'add for Lumira Report

                lumira_request.Add("req_id", request_id.ToString)
                lumira_request.Add("due", http_req_form_duedate)
            Else
                sql_req = "UPDATE requests SET status='CR' WHERE id=" + request_id.ToString
            End If
            lumiraReport.LogRequestReport(request_id, lumira_request)
            dbcomm_req = New OleDbCommand(sql_req, dbconn)

            '////////////////////////////////////////////////////////////
            'CREATE EMAIL AND SEND TO OWNER
            '////////////////////////////////////////////////////////////
            Dim newMail As New MailTemplate
            Dim dueDays As Integer
            Dim dayText As String = "day"
            Dim duereqdate As Date = Date.ParseExact(http_req_form_duedate.Replace(" ", ""), "dd/MMM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None)

            'CALCULATE DUE DAYS
            dueDays = DateDiff(DateInterval.Day, Today.Date, duereqdate.Date)
            If dueDays > 1 Then
                dayText = "days"
            End If

            Dim mail_dict As New Dictionary(Of String, String)
            mail_dict.Add("mail", "CR") 'NEW AI CREATED
            mail_dict.Add("to", userCommon.getMailById(http_req_form_owner))
            mail_dict.Add("{ai_id}", ai_id.ToString)
            mail_dict.Add("{ai_link}", sysConfiguration.getSystemUrl + "sap_ai_view.aspx?id=" + ai_id.ToString)
            mail_dict.Add("{description}", http_req_form_descr) 'MAIL SUBJECT / AI DESCRIPTION
            mail_dict.Add("{duedate}", http_req_form_duedate)
            mail_dict.Add("{accept_link}", sysConfiguration.getSystemUrl + "sap_accept_new_due.aspx?id=" + link.enLink(ai_id.ToString))
            mail_dict.Add("{reject_link}", sysConfiguration.getSystemUrl + "sap_reject_due.aspx?id=" + link.enLink(ai_id.ToString))
            mail_dict.Add("{extension_link}", sysConfiguration.getSystemUrl + "sap_ext.aspx?id=" + link.enLink(ai_id.ToString))
            mail_dict.Add("{need_information}", sysConfiguration.getSystemUrl + "sap_ai_data.aspx?id=" + link.enLink(ai_id.ToString))
            mail_dict.Add("{requestor_name}", userCommon.getNameById(http_req_form_owner))
            mail_dict.Add("{ai_owner}", userCommon.getNameById(http_req_form_owner))
            mail_dict.Add("{app_link}", sysConfiguration.getSystemUrl)
            mail_dict.Add("{contact_mail_link}", "mailto:" & userCommon.getAdminMail & "?subject=Questions about the report")
            mail_dict.Add("{subject}", "AI#" + ai_id.ToString + " Is due in " + dueDays.ToString + " " + dayText)

            newMail.SendNotificationMail(mail_dict)

            '////////////////////////////////////////////////////////////
            'INSERT LOG HERE
            '////////////////////////////////////////////////////////////
            'ai_id, request_id, owner_id, requestor_id, admin_id, event, prev_value, new_value, detail

            'EVENT: AI_CREATED [R1]

            Dim newLog As New Logging

            Dim log_dict As New Dictionary(Of String, String)
            log_dict.Add("ai_id", ai_id.ToString)
            log_dict.Add("request_id", request_id.ToString)
            log_dict.Add("admin_id", userCommon.getId)
            log_dict.Add("owner_id", http_req_form_owner)
            log_dict.Add("event", "AI_CREATED")
            log_dict.Add("detail", "Action Item Created")

            newLog.LogWrite(log_dict)

            Dim redirectTo As String
            redirectTo = sysConfiguration.getSystemUrl + "sap_req.aspx?id=" + http_req_id
            Response.Redirect(redirectTo, False)
        End If

        Dim ro As String = userCommon.getRole()

        '### fabricamos lista de owners ###
        Dim sql_owners As String
        sql_owners = "SELECT * FROM users WHERE role='OW' OR role='AO'"

        dbcomm_ais = New OleDbCommand(sql_owners, dbconn)
        dbread_ais = dbcomm_ais.ExecuteReader()

        While dbread_ais.Read()
            'owner_option.Text = owner_option.Text & "<option value='" & dbread_ais.GetValue(0) & "'>" & dbread_ais.GetValue(1) & "</option>"
            owner.Items.Add(New ListItem(dbread_ais.GetValue(1), dbread_ais.GetValue(0)))
        End While

        Dim extra As String = ""
        If ro = "OW" Then
            extra = " AND owner='" + userCommon.getId() + "'"
        End If
        'ROWS ITERATION
        sql_req = "SELECT * FROM requests WHERE id=" + request_id.ToString
        sql_ais = "SELECT * FROM actionitems WHERE status<>'XX' AND request_id=" + request_id.ToString + extra

        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)

        Dim u As SapUser = New SapUser()
        'ROWS ITERATION
        '#####TODO:#CHECK#IF#REQUEST#EXIST######
        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then

            dbread_req.Read()

            If Not dbread_req.IsDBNull(9) Then
                download_link.HRef = "./downloadmail.ashx?mail=" + dbread_req.GetString(9) + ".eml"
            End If

            req_id.Text = req_id.Text + Convert.ToInt64(dbread_req.GetValue(0)).ToString
            req_detail.Text = dbread_req.GetString(4).Replace("&#39;", "'")
            req_description.Text = dbread_req.GetString(5).Replace("&#39;", "'")
            If Len(req_description.Text) > 250 Then
                req_description.Text = Mid(req_description.Text, 1, 250) & "..."
            End If

            Dim mailHTML As String = dbread_req.GetString(3)
            mail_detail.InnerHtml = HtmlDecode(mailHTML.Replace("&#34;", """").Replace("'", "&#39;"))

            Dim req_created As Date
            Dim req_duedate As Date = actions.requestGetAIsLastDueDate(request_id)
            req_created = dbread_req.GetDateTime(2)
            If req_duedate = Nothing Then
                req_duedate = Today()
            End If

            req_created_year.Text = req_created.Year
            req_created_month.Text = MonthName(req_created.Month, True)
            req_created_day.Text = req_created.Day

            req_duedate_year.Text = req_duedate.Year
            req_duedate_month.Text = MonthName(req_duedate.Month, True)
            req_duedate_day.Text = req_duedate.Day

            Dim tabla As HtmlTable
            tabla = FindControl("ai_list")
            'Dim tabla As Table = Page.FindControl("table#ai-list")

            Dim ai_desc As String

            dbread_ais = dbcomm_ais.ExecuteReader()
            If dbread_ais.HasRows Then
                While dbread_ais.Read()

                    Dim tRow As New HtmlTableRow

                    '<td>151</td>
                    Dim tCell_id As New HtmlTableCell
                    tCell_id.InnerText = Convert.ToInt64(dbread_ais.GetValue(0)).ToString
                    tRow.Cells.Add(tCell_id)

                    '<td class="ai-desc">We need a snack machine to be installed in the new office. I'll be back by 14/02/2015 and I need it ASAP. Please make sure that it has a variety of healthy snacks too. Thanks!</td>
                    Dim tCell_des As New HtmlTableCell
                    ai_desc = dbread_ais.GetString(2).Replace("&#39;", "'")
                    If Len(ai_desc) > 85 Then
                        ai_desc = Mid(ai_desc, 1, 85) & "..."
                    End If
                    tCell_des.InnerText = ai_desc

                    tRow.Cells.Add(tCell_des)

                    Dim un As String = userCommon.getNameById(dbread_ais.GetString(6))
                    '<td>Owner</td>
                    Dim tCell_own As New HtmlTableCell
                    tCell_own.InnerHtml = dbread_ais.GetString(6) + "<br><b>" & un & "</b>"
                    tRow.Cells.Add(tCell_own)

                    '<td>Jan 25, 2015</td>
                    Dim tCell_ctd As New HtmlTableCell
                    Dim ais_created As Date
                    ais_created = dbread_ais.GetDateTime(3)
                    tCell_ctd.InnerHtml = ais_created.Date + "<br><b>" + utils.humanize_Bkw(DateDiff(DateInterval.Day, ais_created.Date, Today.Date)) + "</b>"
                    tRow.Cells.Add(tCell_ctd)

                    '<td>Feb 14, 2015</td>
                    Dim tCell_due As New HtmlTableCell
                    Dim ais_duedate As Date
                    Dim resultHumanizeFwd As String

                    If dbread_ais.IsDBNull(4) Then
                        tCell_due.InnerHtml = "<b>Unset</b>"
                    Else
                        ais_duedate = dbread_ais.GetDateTime(4)
                        resultHumanizeFwd = utils.humanize_Fwd(DateDiff(DateInterval.Day, Today.Date, ais_duedate.Date))
                        tCell_due.InnerHtml = ais_duedate.Date + "<br><b>" + utils.humanize_Fwd(DateDiff(DateInterval.Day, Today.Date, ais_duedate.Date)) + "</b>"

                        If resultHumanizeFwd = "OverDue" Then
                            tCell_due.InnerHtml = ais_duedate.Date + "<br><span style='font-weight: bold;color:rgb(255, 75, 75)'>" + resultHumanizeFwd + "</span>"
                        End If
                    End If
                    tRow.Cells.Add(tCell_due)

                    '<td>Status</td>
                    Dim tCell_sts As New HtmlTableCell
                    tCell_sts.InnerText = utils.ai_Str_Status(dbread_ais.GetString(5))
                    tRow.Cells.Add(tCell_sts)

                    '<td><a href="#">Create</a></td>
                    Dim tCell_btn As New HtmlTableCell

                    If dbread_ais.GetString(5) <> "CP" Then
                        Dim edit_btn As New HtmlAnchor
                        edit_btn.InnerText = "Edit"
                        edit_btn.HRef = sysConfiguration.getSystemUrl + "sap_ai_edit.aspx?id=" + tCell_id.InnerText
                        tCell_btn.Controls.Add(edit_btn)
                    End If

                    If dbread_ais.GetString(5) = "DL" Then
                        'Add Approve button
                        Dim approve_btn As New HtmlAnchor
                        approve_btn.InnerText = "Approve"
                        approve_btn.HRef = sysConfiguration.getSystemUrl + "sap_completed.aspx?id=" + link.enLink(tCell_id.InnerText)
                        tCell_btn.Controls.Add(approve_btn)

                        'Add Reject button
                        Dim reject_btn As New HtmlAnchor
                        reject_btn.InnerText = "Reject"
                        reject_btn.HRef = sysConfiguration.getSystemUrl + "sap_uncompleted.aspx?id=" + link.enLink(tCell_id.InnerText)
                        tCell_btn.Controls.Add(reject_btn)
                    End If

                    tRow.Cells.Add(tCell_btn)

                    tabla.Rows.Add(tRow)

                End While
            Else
                empty_inbox.Text = "empty inbox"
            End If
            dbread_ais.Close()

            min_date.Text = "'" + Now.ToString("yyyy/MM/dd") + "'"
            max_date.Text = "'" + req_duedate.ToString("yyyy/MM/dd") + "'"

        End If

        dbread_req.Close()
        dbconn.Close()

        'SAP ANALYTICS
        REM Dim anal As New SapAnalytics
        'req_timesheet_data.Text = anal.createTimeLine(request_id)
        'section_width.Text = "<style>.timesheet .scale section {width: " + anal.each_section_width.ToString + "px;}</style>"

    End Sub

End Class
