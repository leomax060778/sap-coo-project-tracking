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

            'CLEAR FORM
            'descr.InnerText = ""
            'duedate.Value = ""

            'REQUEST FORM
            If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

                '#####TODO:#CHECK#VALUES################
                Dim http_fullname As String = Request.Form("fullname")
                Dim http_role As String = Request.Form("role")
                Dim http_id As String = Request.Form("id")
				Dim http_email As String = Request.Form("email")
				
                Dim changes As String = ""
                Dim due_updated As Boolean = False
                Dim info_updated As Boolean = False

                'GET DB ACTUAL DATA
                sql_req = "SELECT * FROM users WHERE id='" + http_req_id + "'"
                dbcomm_req = New OleDbCommand(sql_req, dbconn)
                dbread_req = dbcomm_req.ExecuteReader()

                If dbread_req.HasRows Then
                    dbread_req.Read()

                    'IF DB DATA IS DIFFERENT FROM REQUEST THEN
                    If dbread_req.GetString(4) <> http_role Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "role='" + http_role + "'"
                        info_updated = True
                    End If

                    If dbread_req.GetString(1) <> http_fullname Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "name='" + http_fullname + "'"
                        info_updated = True
                    End If
					
					If dbread_req.GetString(2) <> http_email Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "mail='" + http_email + "'"
                        info_updated = True
                    End If

                    'IF THERE ARE CHANGES TO BE DONE
                    If Not String.IsNullOrEmpty(changes) Then
                        sql_req = "UPDATE users SET " + changes + " WHERE id='" + http_req_id + "'"
                        dbcomm_req = New OleDbCommand(sql_req, dbconn)
                        dbcomm_req.ExecuteNonQuery()

                        Dim newLog As New LogSAPTareas

                        changes = ""

                        
                        REM changes = "User Info changed "
                        
                        REM Dim log_dict As New Dictionary(Of String, String)
                        REM 'log_dict.Add("ai_id", ai_id.ToString)
                        REM log_dict.Add("request_id", http_req_id)
                        REM log_dict.Add("admin_id", su.getId)
                        REM 'log_dict.Add("owner_id", "OWNER_ID")
                        REM log_dict.Add("event", "USER_UPDATED")
                        REM log_dict.Add("detail", changes)

                        REM newLog.LogWrite(log_dict)

                    End If

                End If

                Dim redirectTo As String
                redirectTo = syscfg.getSystemUrl + "sap_crud.aspx"
                Response.Redirect(redirectTo, False)
            End If

            Dim extra As String = ""

            'ROWS ITERATION
            sql_req = "SELECT * FROM users WHERE id='" + http_req_id + "'"

            dbcomm_req = New OleDbCommand(sql_req, dbconn)

            'ROWS ITERATION
            '#####TODO:#CHECK#IF#REQUEST#EXIST######
            dbread_req = dbcomm_req.ExecuteReader()
            If dbread_req.HasRows Then

                dbread_req.Read()

                id.Value = dbread_req.GetValue(0).ToString
				fullname.Value = dbread_req.GetValue(1).ToString
                
                role.Value = dbread_req.GetString(4).ToString
				email.Value = dbread_req.GetString(2).ToString
				
				REM req_detail.Text = dbread_req.GetValue(1).ToString
                
                REM Dim mailHTML As String = Convert.ToInt32(dbread_req.GetString(3)).ToString
                REM mail_detail.InnerHtml = HtmlDecode(mailHTML.Replace("&#34;", """").Replace("&#39;", "'"))

                REM descr.InnerText = dbread_req.GetString(5).Replace("&#39;", "'")

                REM Dim req_created As Date
                REM Dim req_duedate As Date
                REM req_created = dbread_req.GetDateTime(2)
                REM If dbread_req.IsDBNull(6) Then
                    REM req_duedate = Today()
                REM Else
                    REM req_duedate = dbread_req.GetDateTime(6)
                REM End If

                REM duedate.Value = req_duedate.ToString("dd/MMM/yyyy") 'req_duedate.Year.ToString + "-" + req_duedate.Month.ToString + "-" + req_duedate.Day.ToString

                link_del_user.HRef = syscfg.getSystemUrl + "sap_user_del.aspx?id=" + id.Value

            End If

            dbread_req.Close()
            dbconn.Close()

            End If

    End Sub

End Class
