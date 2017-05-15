Imports System.Data.OleDb
Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Collections.Generic
Imports LogSAPTareas
Imports MailTemplate
Imports SysConfig
Imports SapActions

Partial Class _Default
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

    Private Sub CommandBtn_Click(ByVal sender As Object, ByVal e As CommandEventArgs)

        Dim syscfg As New SysConfig
        Dim actions As New SapActions

        Dim su as SapUser = new SapUser()
        Dim ro As String = su.getrole()

        If ro <> "OW" And ro <> "AO" Then
            Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
        End If

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
        Dim dbread_ais As OleDbDataReader
        Dim sql, sql_req, sql_ais As String
        Dim ai_new_status, ai_current_status As String
        Dim req_id As Integer

        Dim ai_id As String = CType(e.CommandArgument, String)

        '#####TODO:#CHECK#IF#DB#EXIST###########

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'Get AI current STATUS for VALIDATION
        sql_ais = "SELECT * FROM actionitems WHERE id=" + ai_id
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)
        dbread_ais = dbcomm_ais.ExecuteReader()
        dbread_ais.Read()
        ai_current_status = dbread_ais.GetString(5)
        req_id = dbread_ais.GetInt64(1)

        'Dim comment() As String
        'comment = lnk.CommandArgument.Split(",") ' you can access like comment(0),comment(1)

        ai_new_status = "ERROR"

        'LOG INFORMATION
        Dim log_ai_id As String = ai_id
        Dim log_event As String
        Dim log_detail As String = "Some detail here..."
        Dim log_owner As String = "Current USER here..."
        Dim log_prev_value As String
        Dim log_new_value As String
        Dim log_record As Boolean = False

        Select Case e.CommandName

            Case "Accept"
                If ai_current_status = "PD" Or ai_current_status = "IP" Then

                    ai_new_status = "IP"

                    log_prev_value = ai_current_status
                    log_new_value = ai_new_status
                    log_event = "AI_ACCEPTED"
                    log_detail = "Due Date Accepted"
                    log_record = True

                    actions.requestSetStatus(req_id.ToString, "CR", "IP")

                End If

            Case "Extend"
                If ai_current_status <> "DL" Then
                    Response.Redirect(syscfg.getSystemUrl + "sap_ext.aspx?id=" + ai_id, False)
                End If

            Case "Confirm"
                If ai_current_status <> "PD" And ai_current_status <> "NE" Then

                    ai_new_status = "CF"

                    log_prev_value = ai_current_status
                    log_new_value = ai_new_status
                    log_event = "AI_CONFIRMED"
                    log_detail = "Due Date Confirmed"
                    log_record = True

                End If

            Case "Deliver"
                If ai_current_status <> "PD" And ai_current_status <> "NE" And ai_current_status <> "DL" Then
                    'REDIRECT TO DELIVER PAGE
                    Response.Redirect(syscfg.getSystemUrl + "sap_dlvr.aspx?id=" + ai_id, False)
                End If

            Case Else

                'DO SOMETHING ELSE
                'Message.Text = "Command name not recogized."

        End Select

        If ai_new_status <> "ERROR" Then

            sql_ais = "UPDATE actionitems SET status='" + ai_new_status + "' WHERE id=" + ai_id
            dbcomm_req = New OleDbCommand(sql_ais, dbconn)
            dbcomm_req.ExecuteNonQuery()
            Response.Redirect(syscfg.getSystemUrl + "sap_owner.aspx", False)

            If log_record Then

                Dim newLog As New LogSAPTareas

                Dim log_dict As New Dictionary(Of String, String)
                log_dict.Add("ai_id", log_ai_id)
                log_dict.Add("admin_id", su.getId)
                log_dict.Add("owner_id", log_owner)
                log_dict.Add("prev_value", log_prev_value)
                log_dict.Add("new_value", log_new_value)
                log_dict.Add("event", log_event)
                log_dict.Add("detail", log_detail)

                newLog.LogWrite(log_dict)

            End If

        End If

        dbconn.Close()
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim syscfg As New SysConfig

        Dim su as SapUser = new SapUser()

        Select Case su.getrole()
            Case "AD"
                Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
            Case "guest"
                Response.Redirect(syscfg.getSystemUrl + "sap_new_user.aspx", False)
        End Select

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
        Dim dbread_req, dbread_ais As OleDbDataReader
        Dim sql, sql_req, sql_ais As String

        '#####TODO:#CHECK#IF#DB#EXIST###########

        Dim connString As String = ConfigurationManager.ConnectionStrings("myConnectionString").ConnectionString

        Dim users As SapUser = New SapUser
        current_user.Text = users.getName()

        dbconn = New OleDbConnection(connString)
        dbconn.Open()

        'GET USER
        '#######TODO#########
        'If user is ADMIN redirect to sap_main.aspx
        'If user is OWNER redirect to sap_owner.aspx
        'If user is SUPER redirect to sap_super.aspx

        'HERE GOES OWNER ID NOT REQUEST ID
        '###TODO################################
        Dim request_id As Integer = 2

        'ROWS ITERATION
        'sql_req = "SELECT * FROM requests WHERE id=" + request_id.ToString
        'sql_ais = "SELECT * FROM actionitems WHERE request_id=" + request_id.ToString + " AND owner='" + ru + "'"
        sql_ais = "SELECT * FROM actionitems WHERE owner='" + su.getId() + "' AND status <> 'XX' ORDER BY id DESC"

        'dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)

        'icons
        Dim hot As HtmlImage
        Dim bell As HtmlImage
        Dim ext As HtmlImage
        'hot = New HtmlImage
        'hot.Src = ResolveUrl("~/images/hot.png")
        'bell.Src = ResolveUrl("~/images/bell.png")
        'ext.Src = ResolveUrl("~/images/ext.png")

        'ROWS ITERATION
        '#####TODO:#CHECK#IF#REQUEST#EXIST######
        dbread_ais = dbcomm_ais.ExecuteReader()
        'If dbread_ais.HasRows Then

        'dbread_req.Read()

        req_id.Text = su.getname()

        Dim tabla As HtmlTable
        tabla = FindControl("ai_list")
        'Dim tabla As Table = Page.FindControl("table#ai-list")


        'dbread_ais = dbcomm_ais.ExecuteReader()
        If dbread_ais.HasRows Then
            While dbread_ais.Read()

                Dim ai_id As Long = Convert.ToInt64(dbread_ais.GetValue(0))
                Dim ai_desc As String = dbread_ais.GetString(2)
                Dim ai_created As Date = dbread_ais.GetDateTime(3)
                Dim ai_duedate As Date
                Dim ai_has_duedate As Boolean = Not dbread_ais.IsDBNull(4)
                If ai_has_duedate Then
                    ai_duedate = dbread_ais.GetDateTime(4)
                End If
                Dim ai_status As String = dbread_ais.GetString(5)
                Dim ai_spent_days As Integer = DateDiff(DateInterval.Day, ai_created.Date, Today.Date)
                Dim ai_missing_days As Integer = DateDiff(DateInterval.Day, Today.Date, ai_duedate.Date)

                Dim tRow As New HtmlTableRow

                '<td>151</td>
                Dim tCell_id As New HtmlTableCell
                tCell_id.InnerText = ai_id
                tRow.Cells.Add(tCell_id)

                '<td class="ai-desc">We need a snack machine to be installed in the new office. I'll be back by 14/02/2015 and I need it ASAP. Please make sure that it has a variety of healthy snacks too. Thanks!</td>
                Dim tCell_des As New HtmlTableCell
                if len(ai_desc) > 85 then
                    ai_desc = mid(ai_desc,1,85) & "..."
                End If
                tCell_des.InnerText = ai_desc
                tRow.Cells.Add(tCell_des)

                '<td>Jan 25, 2015</td>
                Dim tCell_ctd As New HtmlTableCell
                tCell_ctd.InnerHtml = ai_created.Date.ToString("dd/MMM/yyyy") + "<br><span class='elapsed'>" + humanize_Bkw(ai_spent_days) + "</span>"
                tRow.Cells.Add(tCell_ctd)

                '<td>Feb 14, 2015</td>
                Dim tCell_due As New HtmlTableCell
                Dim resultHumanizeFwd As String = humanize_Fwd(ai_missing_days)

                If ai_has_duedate Then
                    tCell_due.InnerHtml = ai_duedate.Date.ToString("dd/MMM/yyyy") + "<br><span class='elapsed'>" + humanize_Fwd(ai_missing_days) + "</span>"

                    If resultHumanizeFwd = "Overdue" Then
                        tCell_due.InnerHtml = ai_duedate.Date.ToString("dd/MMM/yyyy") + "<br><span class='elapsed' style='color:rgb(255, 75, 75)'>" + humanize_Fwd(ai_missing_days) + "</span>"
                    End If
                Else
                    tCell_due.InnerHtml = "<b>Unset</b>"
                End If
                If ai_missing_days < 3 And ai_status <> "DL" Then
                    tCell_due.Attributes.Add("class", "hot")
                End If
                tRow.Cells.Add(tCell_due)

                '<td>Status</td>
                Dim tCell_sts As New HtmlTableCell
                tCell_sts.InnerText = ai_Str_Status(ai_status)
                If ai_missing_days < 3 And ai_status <> "DL" Then
                    'tCell_sts.Attributes.Add("class", "hot")
                End If
                'If ai_status = "NE" Then
                '   ext = New HtmlImage
                '   ext.Src = ResolveUrl("~/images/ext.png")
                '   tCell_sts.Controls.Add(ext)
                'End If
                tRow.Cells.Add(tCell_sts)

                '<td><a href="#">Create</a></td>
                Dim tCell_btn As New HtmlTableCell

                If ai_status <> "CF" And ai_status <> "DL" And ai_status <> "NE" And ai_missing_days > 0 And ai_has_duedate Then
                    'Dim ext_btn As New HtmlAnchor
                    'ext_btn.InnerText = "Extend"
                    'ext_btn.HRef = syscfg.getSystemUrl + "sap_ext.aspx?id=" + ai_id.ToString
                    'tCell_btn.Controls.Add(ext_btn)
                    Dim link_ext_btn As New Button
                    link_ext_btn.CommandName = "Extend"
                    link_ext_btn.CommandArgument = ai_id.ToString
                    link_ext_btn.Text = "Extend"
                    'link_ext_btn.Style.Add("float", "right")
                    AddHandler link_ext_btn.Command, AddressOf CommandBtn_Click
                    tCell_btn.Controls.Add(link_ext_btn)
                End If

                'tCell_btn.InnerHtml = "Extend<br>" + tCell_due.InnerHtml

                If ai_status = "PD" And ai_has_duedate Then
                    Dim link_btn As New Button
                    link_btn.CommandName = "Accept"
                    link_btn.CommandArgument = ai_id.ToString
                    link_btn.Text = "Accept"
                    link_btn.Attributes.Add("class", "accept")
                    AddHandler link_btn.Command, AddressOf CommandBtn_Click
                    tCell_btn.Controls.Add(link_btn)
                End If

                If ai_missing_days = 2 And ai_status <> "PD" And ai_status <> "NE" And ai_status <> "CF" Then
                    Dim link_btn As New Button
                    link_btn.CommandName = "Confirm"
                    link_btn.CommandArgument = ai_id.ToString
                    link_btn.Text = "Confirm"
                    AddHandler link_btn.Command, AddressOf CommandBtn_Click
                    tCell_btn.Controls.Add(link_btn)
                End If

                If ai_status <> "PD" And ai_status <> "NE" And ai_status <> "DL" Then 'And ai_missing_days < 1
                    Dim link_btn As New Button
                    link_btn.CommandName = "Deliver"
                    link_btn.CommandArgument = ai_id.ToString
                    link_btn.Text = "Deliver"
                    AddHandler link_btn.Command, AddressOf CommandBtn_Click
                    tCell_btn.Controls.Add(link_btn)
                End If

                tRow.Cells.Add(tCell_btn)

                '<td><a href="request.html">View</a></td>
                'Dim tCell_btn As New HtmlTableCell
                'Dim link_btn As New HtmlAnchor
                'link_btn.HRef = "./sap_req.aspx?id=" + current_req_id.ToString
                'link_btn.InnerText = "View"
                'tCell_btn.Controls.Add(link_btn)
                'tRow.Cells.Add(tCell_btn)

                tRow.Attributes.Add("class", ai_status)
                tRow.Attributes.Add("id", ai_id)

                tabla.Rows.Add(tRow)

            End While
        Else
            empty_inbox.Text = "empty inbox"
        End If
        dbread_ais.Close()
        'End If

        'dbread_req.Close()
        dbconn.Close()
    End Sub

End Class
