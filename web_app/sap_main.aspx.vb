Imports System.Data.OleDb
Imports System.IO
Imports commonLib

Partial Class Default2
    Inherits System.Web.UI.Page

    Dim utilCommon As New commonLib.Utils
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
            Dim log_detail As String = "Some detail here..."
            Dim log_owner As String = "Current USER here..."
            Dim log_record As Boolean = False

            If e.CommandName = "X" Then

                'CHANGE REQUEST STATUS TO CANCELED
                sql_req = "UPDATE requests SET status='XX' WHERE id=" + rq_id
                dbcomm_req = New OleDbCommand(sql_req, dbconn)
                dbcomm_req.ExecuteNonQuery()

                'CHANGE AIS STATUS TO CANCELED
                sql_req = "UPDATE actionitems SET status='XX' WHERE request_id=" + rq_id
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
                log_dict.Add("event", "RQ_CANCELED")
                log_dict.Add("detail", "Request Canceled")

                newLog.LogWrite(log_dict)

                Response.Redirect(sysConfiguration.getSystemUrl + "sap_main.aspx", False)
            End If

            dbconn.Close()

        End If

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Dim user SapUser = New SapUser
        Dim actions As New SapActions

        Select Case userCommon.getRole()
            Case "OW"
                Response.Redirect(sysConfiguration.getSystemUrl + "sap_owner.aspx", False)
            Case "guest"
                Response.Redirect(sysConfiguration.getSystemUrl + "sap_new_user.aspx", False)
                'Case Else
                'Response.Redirect(syscfg.getSystemUrl + "sap_support.aspx", False)
        End Select

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql, sql_req As String

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        'CLEAR FORMS
        new_due_date.Value = ""

        'REQUEST FORM
        If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

            'REQUEST ID
            Dim http_req_id As String = Request("id")
            Dim request_id As Integer

            If Not String.IsNullOrEmpty(http_req_id) And Integer.TryParse(http_req_id, request_id) Then

                '#####TODO:#CHECK#VALUES################
                Dim http_req_form_descr As String = Request.Form("descr")
                Dim http_req_form_owner As String = Request.Form("owner")
                Dim http_req_form_duedate As String = Request.Form("duedate")

                'ID LAST AI CREATED
                Dim ai_id As Long

                'INSERT NEW ROW
                sql = "INSERT INTO actionitems (request_id, description, owner, due, original, status) VALUES (" + http_req_id + ", '" + http_req_form_descr + "', '" + http_req_form_owner + "', '" + http_req_form_duedate + "', '" + http_req_form_duedate + "', 'PD')"
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
                'dbcomm_req.ExecuteNonQuery()

                'NEED ENCRYPTION HERE
                Dim link As New Linker

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
        End If

        Dim filter As String = Request("f")
        Dim extra_where As String = "status <> 'CP' AND status <> 'XX'"
        Dim extra_subq As String = ""
        Select Case filter

            Case "ur"
                extra_where = "ai_count <= 0"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id) AS ai_count"

            Case "nd"
                extra_where = "(status = 'ND' OR ai_count > 0)"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'ND') AS ai_count"

            Case "ap"
                extra_where = "(status = 'PD')"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'PD' AND actionitems.owner IS NOT null) AS ai_count"

            Case "rq"
                extra_where = "(status IN('CR', 'PD') AND ai_count > 0)"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'DL' AND delivery_date IS NOT NULL) AS ai_count"

            Case "du"
                extra_where = "(status <> 'DL' AND status <> 'PD') AND ai_count > 0"
                extra_subq = ", (SELECT COUNT(distinct actionitems.owner) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.owner IS NOT null GROUP BY actionitems.request_id HAVING COUNT(distinct actionitems.owner) > 1 ) AS ai_count"

            Case "ex"
                extra_where = "(status = 'NE') OR (ai_count > 0) "
                extra_subq = ", (SELECT Count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'NE') AS ai_count"

            Case "pr"
                extra_where = "(status = 'DL')"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'DL') AS ai_count"

            Case "dw"
                extra_where = "((status <> 'DL' OR status <> 'PD') AND due IS NOT null AND DATEDIFF(day, TODAY(), due) >= 0 AND DATEDIFF(day, TODAY(), due) <= 7)"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND (actionitems.status <> 'DL' OR actionitems.status <> 'PD') AND (DATEDIFF(day, actionitems.due, TODAY())) >= 0 AND actionitems.owner IS NOT null) AS ai_count"

            Case "ov"
                extra_where = "(status <> 'DL' OR status <> 'PD') AND due IS NOT null AND ai_count >= 0"
                extra_subq = ", (SELECT COUNT(distinct actionitems.id) FROM actionitems WHERE requests.id = actionitems.request_id AND (actionitems.status <> 'DL' AND actionitems.status <> 'PD') AND (DATEDIFF(day, TODAY(), actionitems.due)) < 0 AND actionitems.owner IS NOT null) AS ai_count"

        End Select

        '############ROWS#ITERATION#############
        sql_req = "SELECT * " & extra_subq & " FROM requests WHERE " & extra_where & " ORDER BY id DESC"

        'log("Filter: " + filter + " SQL: " + sql_req)
        'response.write(sql_req)
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        'dbcomm_ais = New OleDbCommand(sql_ais, dbconn)

        '############INSERT#NEW#ROW#############
        'dbcomm.ExecuteScalar()
        current_user.Text = userCommon.getFullName()

        '############ROWS#ITERATION#############
        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then

            Dim ai_desc As String
            Dim tabla As HtmlTable
            tabla = FindControl("ai_list")

            Dim hot As HtmlImage

            While dbread_req.Read()

                Dim tRow As New HtmlTableRow

                tRow.Attributes.Add("class", "req-edit")
                Dim tCell_id As New HtmlTableCell
                Dim current_req_id As Long
                current_req_id = Convert.ToInt64(dbread_req.GetValue(0)) '.ToString
                tCell_id.InnerText = current_req_id
                tRow.Cells.Add(tCell_id)

                Dim tCell_des As New HtmlTableCell
                ai_desc = dbread_req.GetString(5)
                If Len(ai_desc) > 85 Then
                    ai_desc = Mid(ai_desc, 1, 85) & "..."
                End If
                tCell_des.InnerHtml = "<b>" + dbread_req.GetString(4) + "</b><br>" + ai_desc + "<br>"
                tCell_des.Attributes.CssStyle.Add("text-align", "left")
                tRow.Cells.Add(tCell_des)

                Dim req_created As Date
                req_created = dbread_req.GetDateTime(2)
                Dim tCell_ctd As New HtmlTableCell
                Dim req_spent_days As Integer = DateDiff(DateInterval.Day, req_created.Date, Today.Date)
                tCell_ctd.InnerHtml = req_created.Date.ToString("dd/MMM/yyyy") + "<br><span class='elapsed'>" + utilCommon.humanize_Bkw(req_spent_days) + "</span>"

                tRow.Cells.Add(tCell_ctd)

                '<td>Feb 14, 2015</td>
                'FIND FIRST DUEDATE IN AI LIST
                '###TODO#######################
                Dim req_duedate As Date
                '##############################################
                '##############################################
                '##############################################
                '##############################################
                '##############################################
                '##############################################
                '##############################################
                Dim tCell_due As New HtmlTableCell
                If Not dbread_req.IsDBNull(6) Then
                    req_duedate = dbread_req.GetDateTime(6)
                Else
                    req_duedate = Nothing
                End If

                'req_duedate = 'actions.requestGetAIsLastDueDate(current_req_id)
                Dim req_missing_days As Integer = 0
                If req_duedate = Nothing Then
                    Dim new_due_btn As New HtmlAnchor

                    req_missing_days = -1
                    req_duedate = Today()

                    new_due_btn.InnerText = "Set Due Date"
                    new_due_btn.HRef = sysConfiguration.getSystemUrl + "sap_req_edit.aspx?id=" + current_req_id.ToString
                    tCell_due.Controls.Add(new_due_btn)
                Else
                    'req_duedate = dbread_req.GetDateTime(6)
                    req_missing_days = DateDiff(DateInterval.Day, Today.Date, req_duedate.Date)
                    tCell_due.InnerHtml = req_duedate.Date.ToString("dd/MMM/yyyy") + "<br><span class='elapsed' style='color:rgb(255, 75, 75)'>" + utilCommon.humanize_Fwd(req_missing_days) + "</span>"
                End If

                tRow.Cells.Add(tCell_due)

                '<td>Status</td>
                Dim tCell_ais As New HtmlTableCell
                Dim req_status As String = dbread_req.GetString(7)
                'tCell_ais.InnerHtml = "<div style='float: left;margin-top: 5px;'>" & syscfg.req_Str_Status(req_status) & "</div>"
                tCell_ais.InnerHtml = utilCommon.ai_Str_Status(req_status)
                If req_status = "00" Then
                    hot = New HtmlImage
                    hot.Src = ResolveUrl("~/images/hot.png")
                    hot.Attributes.CssStyle.Add("height", "20px")
                    hot.Attributes.CssStyle.Add("float", "left")
                    hot.Attributes.CssStyle.Add("margin-right", "5px")
                    tCell_ais.Controls.Add(hot)
                End If
                tRow.Cells.Add(tCell_ais)

                '<td><a href="request.html">View</a></td>
                Dim tCell_btn As New HtmlTableCell

                Dim text_btn As String
                If req_status = "PD" Then
                    Dim nd_btn As New HtmlAnchor
                    nd_btn.HRef = sysConfiguration.getSystemUrl + "sap_data.aspx?id=" + current_req_id.ToString
                    nd_btn.InnerText = "?"
                    tCell_btn.Controls.Add(nd_btn)
                    tRow.Cells.Add(tCell_btn)
                    'Else
                    '   text_btn = "Update"
                End If

                tRow.Attributes.Add("id", current_req_id.ToString)
                current_url.HRef = sysConfiguration.getSystemUrl

                'Dim link_btn As New HtmlAnchor
                'link_btn.HRef = syscfg.getSystemUrl + "sap_req.aspx?id=" + current_req_id.ToString
                'link_btn.InnerText = "Create"
                'tCell_btn.Controls.Add(link_btn)

                Dim edit_btn As New HtmlAnchor
                'Dim element As New document.createElement("img")
                'element.setAttribute("src", "./images/edit.png")
                'edit_btn.appendChild(element)
                'Dim pencil As New HtmlImage
                'pencil.Attributes.Add("src", "./images/edit.png")
                edit_btn.InnerText = "Edit"
                'edit_btn.InnerHtml = "<img src='./images/edit.png'>"
                edit_btn.HRef = sysConfiguration.getSystemUrl + "sap_req_edit.aspx?id=" + current_req_id.ToString
                tCell_btn.Controls.Add(edit_btn)

                tRow.Cells.Add(tCell_btn)

                If filter <> "od" Or (filter = "od" And req_missing_days < 7 And req_missing_days > -1) Then
                    tabla.Rows.Add(tRow)
                End If

            End While
        Else
            empty_inbox.Text = "empty inbox"
        End If

        dbread_req.Close()
        dbconn.Close()
    End Sub

    Sub log(ByVal message As String)
        Dim strFile As String = "d:\webapps\test\log.txt"
        Dim fileExists As Boolean = File.Exists(strFile)
        Using sw As New StreamWriter(File.Open(strFile, FileMode.Append))
            sw.WriteLine("[" & DateTime.Now & "] " & message)
        End Using
    End Sub

End Class
