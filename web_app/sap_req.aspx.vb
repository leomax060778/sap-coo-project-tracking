Imports System.Data.OleDb
Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.HttpUtility
Imports System.Collections.Generic
Imports SysConfig
Imports SapActions
Imports SapAnalytics
Imports Linker
'Imports MailTemplate

Partial Class _Default
    Inherits System.Web.UI.Page

    Private Sub CommandBtn_Click(ByVal sender As Object, ByVal e As CommandEventArgs)

        Dim syscfg As New SysConfig
        Dim actions As New SapActions

        Dim users As SapUser = New SapUser
        Dim ro As String = users.getRole()

        If ro = "OW" Then
            Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
        Else
            Dim dbconn As OleDbConnection
            Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
            Dim dbread_ais As OleDbDataReader
            Dim sql, sql_req, sql_ais As String
            Dim ai_new_status, ai_current_status As String

            Dim rq_id As String = CType(e.CommandArgument, String)

            '#####TODO:#CHECK#IF#DB#EXIST###########

            dbconn = New OleDbConnection(syscfg.getConnection)
            dbconn.Open()

            'LOG INFORMATION
            Dim log_rq_id As String = rq_id
            Dim log_event As String
            Dim log_detail As String = "Some detail here..."
            Dim log_owner As String = "Current USER here..."
            Dim log_prev_value As String
            Dim log_new_value As String
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

                Dim newLog As New LogSAPTareas

                Dim log_dict As New Dictionary(Of String, String)
                log_dict.Add("request_id", rq_id)
                log_dict.Add("admin_id", users.getId)
                log_dict.Add("event", "AI_CANCELED")
                log_dict.Add("detail", "Action Item Canceled")

                newLog.LogWrite(log_dict)

                Response.Redirect(syscfg.getSystemUrl + "sap_req.aspx?id=" + Request("id"), False)
            End If

            dbconn.Close()

        End If

    End Sub

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
            Case "CP"
                result = "Completed"
            Case Else
                result = "Unset"
        End Select
        Return result
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim su as SapUser = new SapUser()
        Dim link As New Linker

        current_user.Text = su.getName()

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
        Dim dbread_req, dbread_ais As OleDbDataReader
        Dim sql, sql_req, sql_ais As String

        '#####TODO:#CHECK#IF#DB#EXIST###########

        Dim syscfg As New SysConfig
        Dim actions As New SapActions

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

            'IF THE RQ DUE DATE IS UNSET THEN SET IT TO THE AI DUE DATE
            'AND CHANGE REQUEST STATUS TO CREATED
            If actions.requestIsUnset(http_req_id) Then
                sql_req = "UPDATE requests SET status='CR', due='" + http_req_form_duedate + "' WHERE id=" + request_id.ToString
            Else
                sql_req = "UPDATE requests SET status='CR' WHERE id=" + request_id.ToString
            End If
            dbcomm_req = New OleDbCommand(sql_req, dbconn)
            'bdbcomm_req.ExecuteNonQuery()

            'NEED ENCRYPTION HERE
            'Dim link As New Linker

            '////////////////////////////////////////////////////////////
            'CREATE EMAIL AND SEND TO OWNER
            '////////////////////////////////////////////////////////////
            Dim newMail As New MailTemplate

            Dim mail_dict As New Dictionary(Of String, String)
            mail_dict.Add("mail", "CR") 'NEW AI CREATED
            mail_dict.Add("to", su.getMailById(http_req_form_owner))
            mail_dict.Add("{ai_id}", ai_id.ToString)
            mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + ai_id.ToString)
            mail_dict.Add("{description}", http_req_form_descr) 'MAIL SUBJECT / AI DESCRIPTION
            mail_dict.Add("{duedate}", http_req_form_duedate)
            mail_dict.Add("{accept_link}", syscfg.getSystemUrl + "sap_accept_new_due.aspx?id=" + link.enLink(ai_id.ToString))
			mail_dict.Add("{reject_link}", syscfg.getSystemUrl + "sap_reject_due.aspx?id=" + link.enLink(ai_id.ToString))
            mail_dict.Add("{extension_link}", syscfg.getSystemUrl + "sap_ext.aspx?id=" + link.enLink(ai_id.ToString))
            mail_dict.Add("{requestor_name}", su.getNameById(http_req_form_owner))
            mail_dict.Add("{ai_owner}", su.getNameById(http_req_form_owner))
            mail_dict.Add("{app_link}", syscfg.getSystemUrl)
            mail_dict.Add("{contact_mail_link}", "mailto:" & su.getAdminMail & "?subject=Questions about the report")

            newMail.SendNotificationMail(mail_dict)

            '////////////////////////////////////////////////////////////
            'INSERT LOG HERE
            '////////////////////////////////////////////////////////////
            'ai_id, request_id, owner_id, requestor_id, admin_id, event, prev_value, new_value, detail

            'EVENT: AI_CREATED [R1]

            Dim newLog As New LogSAPTareas

            Dim log_dict As New Dictionary(Of String, String)
            log_dict.Add("ai_id", ai_id.ToString)
            log_dict.Add("request_id", request_id.ToString)
            log_dict.Add("admin_id", su.getId)
            log_dict.Add("owner_id", http_req_form_owner)
            log_dict.Add("event", "AI_CREATED")
            log_dict.Add("detail", "Action Item Created")

            newLog.LogWrite(log_dict)

            Dim redirectTo As String
            redirectTo = syscfg.getSystemUrl + "sap_req.aspx?id=" + http_req_id
            Response.Redirect(redirectTo, False)
        End If

        Dim ro As String = su.getRole()

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
            extra = " AND owner='" + su.getId() + "'"
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

                    Dim un As String = u.getNameById(dbread_ais.GetString(6))
                    '<td>Owner</td>
                    Dim tCell_own As New HtmlTableCell
                    tCell_own.InnerHtml = dbread_ais.GetString(6) + "<br><b>" & un & "</b>"
                    tRow.Cells.Add(tCell_own)

                    '<td>Jan 25, 2015</td>
                    Dim tCell_ctd As New HtmlTableCell
                    Dim ais_created As Date
                    ais_created = dbread_ais.GetDateTime(3)
                    tCell_ctd.InnerHtml = ais_created.Date + "<br><b>" + HumanizeBkw(DateDiff(DateInterval.Day, ais_created.Date, Today.Date)) + "</b>"
                    tRow.Cells.Add(tCell_ctd)

                    '<td>Feb 14, 2015</td>
                    Dim tCell_due As New HtmlTableCell
                    Dim ais_duedate As Date
                    Dim resultHumanizeFwd As String

                    If dbread_ais.IsDBNull(4) Then
                        tCell_due.InnerHtml = "<b>Unset</b>"
                    Else
                        ais_duedate = dbread_ais.GetDateTime(4)
                        resultHumanizeFwd = HumanizeFwd(DateDiff(DateInterval.Day, Today.Date, ais_duedate.Date))
                        tCell_due.InnerHtml = ais_duedate.Date + "<br><b>" + HumanizeFwd(DateDiff(DateInterval.Day, Today.Date, ais_duedate.Date)) + "</b>"

                        If resultHumanizeFwd = "OverDue" Then
                            tCell_due.InnerHtml = ais_duedate.Date + "<br><span style='font-weight: bold;color:rgb(255, 75, 75)'>" + resultHumanizeFwd + "</span>"
                        End If
                    End If
                    tRow.Cells.Add(tCell_due)

                    '<td>Status</td>
                    Dim tCell_sts As New HtmlTableCell
                    tCell_sts.InnerText = ai_Str_Status(dbread_ais.GetString(5))
                    tRow.Cells.Add(tCell_sts)

                    '<td><a href="#">Create</a></td>
                    Dim tCell_btn As New HtmlTableCell

                    If dbread_ais.GetString(5) <> "CP" Then
                        Dim edit_btn As New HtmlAnchor
                        edit_btn.InnerText = "Edit"
                        edit_btn.HRef = syscfg.getSystemUrl + "sap_ai_edit.aspx?id=" + tCell_id.InnerText
                        tCell_btn.Controls.Add(edit_btn)
                    End If

                    If dbread_ais.GetString(5) = "DL" Then
                        'Add Approve button
                        Dim approve_btn As New HtmlAnchor
                        approve_btn.InnerText = "Approve"
                        approve_btn.HRef = syscfg.getSystemUrl + "sap_completed.aspx?id=" + link.enLink(tCell_id.InnerText)
                        tCell_btn.Controls.Add(approve_btn)

                        'Add Reject button
                        Dim reject_btn As New HtmlAnchor
                        reject_btn.InnerText = "Reject"
                        reject_btn.HRef = syscfg.getSystemUrl + "sap_uncompleted.aspx?id=" + link.enLink(tCell_id.InnerText)
                        tCell_btn.Controls.Add(reject_btn)
                    End If

                    tRow.Cells.Add(tCell_btn)

                    tabla.Rows.Add(tRow)

                End While
            Else
                empty_inbox.Text = "empty inbox"
            End If
            dbread_ais.Close()

            min_date.Text = "'" + now.ToString("yyyy/MM/dd") + "'"
            max_date.Text = "'" + req_duedate.ToString("yyyy/MM/dd") + "'"

        End If

        dbread_req.Close()
        dbconn.Close()

        'SAP ANALYTICS
        Dim anal As New SapAnalytics
        'req_timesheet_data.Text = anal.createTimeLine(request_id)
        'section_width.Text = "<style>.timesheet .scale section {width: " + anal.each_section_width.ToString + "px;}</style>"

    End Sub

End Class
