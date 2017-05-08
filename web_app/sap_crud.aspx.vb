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
            Case Else
                'Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
        End Select

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
        Dim dbread_req, dbread_ais As OleDbDataReader
        Dim sql, sql_req, sql_ais As String

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'CLEAR FORMS
        'new_due_date.Value = ""

        'REQUEST FORM
        If Request.ServerVariables("REQUEST_METHOD") = "POST" Then


            '#####TODO:#CHECK#VALUES################
            Dim http_req_user_id As String = Request.Form("user_id")
            Dim http_req_user_name As String = Request.Form("user_name")
            Dim http_req_user_role As String = Request.Form("user_role")
            Dim http_req_user_email As String = Request.Form("user_email")

            'IF USER DOES NOT EXIST IS A GUEST
            If users.getRoleById(http_req_user_id) = "guest" Then

                'INSERT NEW ROW
                sql = "INSERT INTO users (id, name, mail, role) VALUES ('" + http_req_user_id + "', '" + http_req_user_name + "', '" + http_req_user_email + "', '" + http_req_user_role + "')"
                dbcomm = New OleDbCommand(sql, dbconn)
                dbcomm.ExecuteNonQuery()
                'dbcomm.CommandText = "SELECT @@IDENTITY"
                'dbcomm.ExecuteScalar()

                '////////////////////////////////////////////////////////////
                'INSERT LOG HERE
                '////////////////////////////////////////////////////////////
                'ai_id, request_id, owner_id, requestor_id, admin_id, event, prev_value, new_value, detail

                'EVENT: AI_CREATED [R1]

                Dim newLog As New LogSAPTareas

                Dim log_dict As New Dictionary(Of String, String)
                log_dict.Add("admin_id", users.getId)
                log_dict.Add("event", "USER_CREATED")
                log_dict.Add("detail", http_req_user_id & " - " & http_req_user_name & " - " & http_req_user_role & " - " & http_req_user_email)

                newLog.LogWrite(log_dict)

            End If

            Response.Redirect(syscfg.getSystemUrl + "sap_crud.aspx", False)

        End If

        '############INSERT#NEW#ROW#############
        'sql = "INSERT INTO requests (requestor, mail, subject, detail, due) VALUES ('pablo@sap.com', 'Dear, Juan Perez. Need a new slaptron for next tuesday. Thanks!', 'New Slaptron', 'Next tuesday', '2015-02-25')"

        '############ROWS#ITERATION#############
        sql_req = "SELECT * FROM users;"
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

                Dim tCell_id As New HtmlTableCell
                Dim current_req_id As String
                current_req_id = dbread_req.GetString(0) '.ToString
                tCell_id.InnerText = current_req_id
                tRow.Cells.Add(tCell_id)

                Dim tCell_des As New HtmlTableCell
                ai_desc = dbread_req.GetString(1)
                tCell_des.InnerHtml = ai_desc
                'tCell_des.Attributes.CssStyle.Add("text-align", "left")
                tRow.Cells.Add(tCell_des)

                Dim req_created As String
                req_created = dbread_req.GetString(2)
                Dim tCell_ctd As New HtmlTableCell
                tCell_ctd.InnerHtml = req_created
                tRow.Cells.Add(tCell_ctd)

                'Dim req_status As String = dbread_req.GetBoolean(3).ToString
                'Dim tCell_due As New HtmlTableCell
                'tCell_due.InnerHtml = req_status
                'tRow.Cells.Add(tCell_due)

                Dim req_role As String = dbread_req.GetString(4)
                Dim tCell_role As New HtmlTableCell
                Select Case req_role
                    Case "OW"
                        req_role = "Owner"
                    Case "RQ"
                        req_role = "Requestor"
                    Case "AD"
                        req_role = "Admin"
                End Select
                tCell_role.InnerHtml = req_role
                tRow.Cells.Add(tCell_role)

                '<td><a href="request.html">View</a></td>
                Dim tCell_btn As New HtmlTableCell

                tRow.Attributes.Add("id", current_req_id.ToString)
                current_url.HRef = syscfg.getSystemUrl

                Dim edit_btn As New HtmlAnchor
                edit_btn.InnerText = "Edit"
                edit_btn.HRef = syscfg.getSystemUrl + "sap_user_edit.aspx?id=" + current_req_id.ToString
                tCell_btn.Controls.Add(edit_btn)

                tRow.Cells.Add(tCell_btn)

                tabla.Rows.Add(tRow)

            End While
        Else
            empty_inbox.Text = "empty list"
        End If

        dbread_req.Close()
        dbconn.Close()
    End Sub

End Class
