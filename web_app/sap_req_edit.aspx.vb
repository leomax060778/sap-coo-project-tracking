Imports System.Data.OleDb
Imports System.Web.HttpUtility
Imports commonLib

Partial Class _Default
    Inherits System.Web.UI.Page

    Dim sysConfiguration As New SystemConfiguration
    Dim userCommon As New commonLib.SapUser

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim ro As String = userCommon.getRole()

        Dim actions As New SapActions

        current_user.Text = userCommon.getFullName()

        If ro = "OW" Then
            Response.Redirect(sysConfiguration.getSystemUrl + "sap_main.aspx", False)
        Else

            Dim dbconn As OleDbConnection
            Dim dbcomm_req As OleDbCommand
            Dim dbread_req As OleDbDataReader
            Dim sql_req As String

            '#####TODO:#CHECK#IF#DB#EXIST###########
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

            'REQUEST FORM
            If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

                '#####TODO:#CHECK#VALUES################
                Dim http_req_form_subj As String = Request.Form("subj")
                Dim http_req_form_descr As String = Request.Form("descr")
                Dim http_req_form_duedate As String = Request.Form("duedate")
                Dim changes As String = ""
                Dim due_updated As Boolean = False
                Dim info_updated As Boolean = False

                'Create a Lumira Request
                Dim lumiraReport As New LumiraReports
                Dim lumira_request As New Dictionary(Of String, String)

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
                        lumira_request.Add("subject", http_req_form_subj.Replace("&#34;", """").Replace("'", "&#39;"))
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

                    Dim duereqdate As Date = Date.ParseExact(http_req_form_duedate.Replace(" ", ""), "dd/MMM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None)

                    If duedb.Year <> duereqdate.Year Or duedb.Month <> duereqdate.Month Or duedb.Day <> duereqdate.Day Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "due='" + duereqdate.Year.ToString + "-" + duereqdate.Month.ToString + "-" + duereqdate.Day.ToString + "'"
                        lumira_request.Add("due", duereqdate.ToString("yyyy/MM/dd HH:mm:ss"))
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
                            lumira_request.Add("created", Date.UtcNow.ToString("yyyy/MM/dd HH:mm:ss"))
                        End If

                        Dim newLog As New Logging
                        Dim log_dict As New Dictionary(Of String, String)

                        changes = ""

                        If due_updated Then
                            changes = "Info changed "
                        End If

                        If info_updated Then
                            changes = changes + "Due Date changed"
                        End If

                        'Add log entries to dictionary
                        log_dict.Add("request_id", request_id.ToString)
                        log_dict.Add("admin_id", userCommon.getId)
                        log_dict.Add("event", "RQ_UPDATED")
                        log_dict.Add("detail", changes)

                        'Insert entry for Lumira Request
                        lumiraReport.LogRequestReport(request_id, lumira_request)

                        newLog.LogWrite(log_dict)

                    End If

                End If

                Dim redirectTo As String
                redirectTo = sysConfiguration.getSystemUrl + "sap_req.aspx?id=" + http_req_id
                Response.Redirect(redirectTo, False)
            End If

            Dim extra As String = ""

            'ROWS ITERATION
            '#####TODO:#CHECK#IF#REQUEST#EXIST######
            sql_req = "SELECT * FROM requests WHERE id=" + request_id.ToString
            dbcomm_req = New OleDbCommand(sql_req, dbconn)
            dbread_req = dbcomm_req.ExecuteReader()

            If dbread_req.HasRows Then

                dbread_req.Read()

                req_id.Text = req_id.Text + Convert.ToInt64(dbread_req.GetValue(0)).ToString

                req_detail.Text = dbread_req.GetString(4).Replace("&#39;", "'")
                subj.InnerText = req_detail.Text

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

                duedate.Value = req_duedate.ToString("dd/MMM/yyyy")
                link_del_req.HRef = sysConfiguration.getSystemUrl + "sap_req_del.aspx?id=" + Convert.ToInt64(dbread_req.GetValue(0)).ToString

            End If

            dbread_req.Close()
            dbconn.Close()

        End If

    End Sub

End Class
