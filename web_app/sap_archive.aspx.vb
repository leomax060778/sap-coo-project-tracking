Imports System.Data.OleDb
Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports SysConfig

Partial Class Default2
    Inherits System.Web.UI.Page

    Private Function humanize_Fwd(ByVal i As Integer) As String
        Dim result As String
        Select Case i
            Case Is < 0
                result = "Overdue"
            Case 0
                result = "Today"
            Case 1
                result = "Tomorrow"
            Case Else
                result = "within " + i.ToString + " days"
        End Select
        Return result
    End Function

    Private Function humanize_Bkw(ByVal i As Integer) As String
        Dim result As String
        Select Case i
            Case Is < 0
                result = "Error"
            Case 0
                result = "Today"
            Case 1
                result = "Yesterday"
            Case Else
                result = i.ToString + " days ago"
        End Select
        Return result
    End Function

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

                Dim newLog As New LogSAPTareas

                Dim log_dict As New Dictionary(Of String, String)
                log_dict.Add("request_id", rq_id)
                log_dict.Add("admin_id", users.getId)
                log_dict.Add("event", "RQ_CANCELED")
                log_dict.Add("detail", "Request Canceled")

                newLog.LogWrite(log_dict)

                Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
            End If

            dbconn.Close()

        End If

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim syscfg As New SysConfig
        Dim users As SapUser = New SapUser
        Dim actions As New SapActions

        Select Case users.getRole()
            Case "OW"
                Response.Redirect(syscfg.getSystemUrl + "sap_owner.aspx", False)
            Case "guest"
                Response.Redirect(syscfg.getSystemUrl + "sap_new_user.aspx", False)
                'Case Else
                'Response.Redirect(syscfg.getSystemUrl + "sap_support.aspx", False)
        End Select

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
        Dim dbread_req, dbread_ais As OleDbDataReader
        Dim sql, sql_req, sql_ais As String

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'CLEAR FORMS
        new_due_date.Value = ""



        '############INSERT#NEW#ROW#############
        'sql = "INSERT INTO requests (requestor, mail, subject, detail, due) VALUES ('pablo@sap.com', 'Dear, Juan Perez. Need a new slaptron for next tuesday. Thanks!', 'New Slaptron', 'Next tuesday', '2015-02-25')"

        '############ROWS#ITERATION#############
        sql_req = "SELECT * FROM requests WHERE status = 'CP' ORDER BY requested DESC" 'OR status = 'XX'
        'sql_ais = "SELECT * FROM actionitems WHERE request_id=" + request_id.ToString

        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        'dbcomm_ais = New OleDbCommand(sql_ais, dbconn)

        '############INSERT#NEW#ROW#############
        'dbcomm.ExecuteScalar()

        current_user.Text = users.getName()



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

                '<td>134</td>
                Dim tCell_id As New HtmlTableCell
                Dim current_req_id As Long
                current_req_id = Convert.ToInt64(dbread_req.GetValue(0)) '.ToString
                tCell_id.InnerText = current_req_id
                tRow.Cells.Add(tCell_id)

                '<td class="ai-desc">We need a snack machine to be installed in the new office. I'll be back by 14/02/2015 and I need it ASAP. Please make sure that it has a variety of healthy snacks too. Thanks!</td>
                Dim tCell_des As New HtmlTableCell
                ai_desc = dbread_req.GetString(5)
                If Len(ai_desc) > 85 Then
                    ai_desc = Mid(ai_desc, 1, 85) & "..."
                End If
                tCell_des.InnerHtml = "<b>" + dbread_req.GetString(4) + "</b><br>" + ai_desc + "<br>"
                tCell_des.Attributes.CssStyle.Add("text-align", "left")
                tRow.Cells.Add(tCell_des)

                '<td>Jan 25, 2015</td>
                Dim req_created As Date
                req_created = dbread_req.GetDateTime(2)
                Dim tCell_ctd As New HtmlTableCell
                Dim req_spent_days As Integer = DateDiff(DateInterval.Day, req_created.Date, Today.Date)
                tCell_ctd.InnerHtml = req_created.Date.ToString("dd/MMM/yyyy") + "<br><span class='elapsed'>" + humanize_Bkw(req_spent_days) + "</span>"

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
                req_duedate = actions.requestGetAIsLastDueDate(current_req_id)
                If req_duedate = Nothing Then

                    tCell_due.InnerHtml = "Not Set"

                Else
                    'req_duedate = dbread_req.GetDateTime(6)
                    Dim req_missing_days As Integer = DateDiff(DateInterval.Day, Today.Date, req_duedate.Date)
                    tCell_due.InnerHtml = req_duedate.Date.ToString("dd/MMM/yyyy") + "<br><span class='elapsed'>" + humanize_Fwd(req_missing_days) + "</span>"

                End If

                tRow.Cells.Add(tCell_due)

                '<td>Status</td>
                Dim tCell_ais As New HtmlTableCell
                Dim req_status As String = dbread_req.GetString(7)
                'tCell_ais.InnerHtml = "<div style='float: left;margin-top: 5px;'>" & syscfg.req_Str_Status(req_status) & "</div>"
                tCell_ais.InnerHtml = syscfg.req_Str_Status(req_status)
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
                    nd_btn.HRef = syscfg.getSystemUrl + "sap_data.aspx?id=" + current_req_id.ToString
                    nd_btn.InnerText = "?"
                    tCell_btn.Controls.Add(nd_btn)
                    tRow.Cells.Add(tCell_btn)
                    'Else
                    '   text_btn = "Update"
                End If

                tRow.Attributes.Add("id", current_req_id.ToString)
                current_url.HRef = syscfg.getSystemUrl

                'Dim link_btn As New HtmlAnchor
                'link_btn.HRef = syscfg.getSystemUrl + "sap_req.aspx?id=" + current_req_id.ToString
                'link_btn.InnerText = "Create"
                'tCell_btn.Controls.Add(link_btn)

                'Dim edit_btn As New HtmlAnchor
                'Dim element As New document.createElement("img")
                'element.setAttribute("src", "./images/edit.png")
                'edit_btn.appendChild(element)
                'Dim pencil As New HtmlImage
                'pencil.Attributes.Add("src", "./images/edit.png")
                'edit_btn.InnerText = "Edit"
                'edit_btn.InnerHtml = "<img src='./images/edit.png'>"
                'edit_btn.HRef = syscfg.getSystemUrl + "sap_req_edit.aspx?id=" + current_req_id.ToString
                'tCell_btn.Controls.Add(edit_btn)

                'tRow.Cells.Add(tCell_btn)

                tabla.Rows.Add(tRow)

            End While
        Else
            empty_inbox.Text = "empty inbox"
        End If

        dbread_req.Close()
        dbconn.Close()
    End Sub

End Class
