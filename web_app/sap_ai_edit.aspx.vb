Imports System.Data.OleDb
Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.HttpUtility
Imports System.Collections.Generic
Imports SysConfig
Imports SapActions
Imports Linker
'Imports MailTemplate

'#############################################################################################################
'#############################################################################################################
'NOT IMPLEMENTED / TO DO
'#############################################################################################################
'#############################################################################################################

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
        Dim link As New Linker

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
                Dim http_req_form_owner As String = Request.Form("owner")
                Dim http_req_form_descr As String = Request.Form("descr")
                Dim http_req_form_status As String = Request.Form("status")
                Dim http_req_form_duedate As String = Request.Form("duedate")
                Dim changes As String = ""
                Dim redirect_req As Integer
                Dim sendMail As Boolean = False
                Dim currentOwner As String
                Dim currentStatus As String
                Dim currentDesc As String

                'GET DB ACTUAL DATA
                sql_req = "SELECT * FROM actionitems WHERE id=" + request_id.ToString
                dbcomm_req = New OleDbCommand(sql_req, dbconn)
                dbread_req = dbcomm_req.ExecuteReader()

                If dbread_req.HasRows Then
                    dbread_req.Read()

                    redirect_req = dbread_req.GetInt64(1)

                    'IF DB DATA IS DIFFERENT FROM REQUEST THEN
                    currentOwner = dbread_req.GetString(6)
                    If currentOwner <> http_req_form_owner Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "owner='" + http_req_form_owner + "'"
                        sendMail = True
                    End If

                    currentStatus = dbread_req.GetString(5)

                    currentDesc = dbread_req.GetString(2)
                    If currentDesc <> http_req_form_descr Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "description='" + http_req_form_descr.Replace("&#34;", """").Replace("'", "&#39;") + "'"
                        sendMail = True
                    End If

                    Dim duedb As Date

                    If Not String.IsNullOrEmpty(http_req_form_duedate) Then

                        If Not dbread_req.IsDBNull(4) Then
                            duedb = dbread_req.GetDateTime(4)
                        Else
                            duedb = Today.Date
                        End If

                        Dim duereqdate As Date = Date.ParseExact(http_req_form_duedate.Replace(" ", ""), "dd/MMM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None)
                        'Date("#" & http_req_form_duedate.Replace(" ", "") & "#")
                        'Dim duereq() As String = http_req_form_duedate.Split("-")
                        If duedb.Year <> duereqdate.Year Or duedb.Month <> duereqdate.Month Or duedb.Day <> duereqdate.Day Then
                            If Not String.IsNullOrEmpty(changes) Then
                                changes = changes + ", "
                            End If
                            changes = changes + "due='" + duereqdate.Year.ToString + "-" + duereqdate.Month.ToString + "-" + duereqdate.Day.ToString + "'"
                            sendMail = True
                        End If

                    End If

                    'IF THERE ARE CHANGES TO BE DONE
                    If String.IsNullOrEmpty(changes) Then

                        If currentStatus <> http_req_form_status and Not sendMail Then
                            changes = "status='" + http_req_form_status + "'"
                        End If
                    Else
                        changes = changes + ",status='PD'"
                    End If

                    If Not String.IsNullOrEmpty(changes) Then
                        sql_req = "UPDATE actionitems SET " + changes + " WHERE id=" + request_id.ToString
                        dbcomm_req = New OleDbCommand(sql_req, dbconn)
                        dbcomm_req.ExecuteNonQuery()

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
                        Dim newMail As New MailTemplate

                        Dim mail_dict As New Dictionary(Of String, String)
                        mail_dict.Add("mail", "CR") 'NEW AI CREATED
                        mail_dict.Add("to", su.getMailById(http_req_form_owner))
                        mail_dict.Add("{ai_id}", request_id.ToString)
                        mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + request_id.ToString)
                        mail_dict.Add("{ai_owner}", su.getNameById(http_req_form_owner))
                        mail_dict.Add("{description}", http_req_form_descr) 'MAIL SUBJECT / AI DESCRIPTION
                        mail_dict.Add("{duedate}", http_req_form_duedate)
                        mail_dict.Add("{accept_link}", syscfg.getSystemUrl + "sap_accept_new_due.aspx?id=" + link.enLink(request_id.ToString))
						mail_dict.Add("{reject_link}", syscfg.getSystemUrl + "sap_reject_due.aspx?id=" + link.enLink(request_id.ToString))
                        mail_dict.Add("{extension_link}", syscfg.getSystemUrl + "sap_ext.aspx?id=" + link.enLink(request_id.ToString))
                        mail_dict.Add("{requestor_name}", su.getNameById(http_req_form_owner))
                        mail_dict.Add("{app_link}", syscfg.getSystemUrl)
                        mail_dict.Add("{contact_mail_link}", "mailto:" & su.getAdminMail & "?subject=Questions about the report")


                        If sendMail Then
                            newMail.SendNotificationMail(mail_dict)
                        End If

                        '////////////////////////////////////////////////////////////
                        'INSERT LOG HERE
                        '////////////////////////////////////////////////////////////
                        'ai_id, request_id, owner_id, requestor_id, admin_id, event, prev_value, new_value, detail

                        'EVENT: AI_CREATED [R1]

                        Dim newLog As New LogSAPTareas

                        Dim log_dict As New Dictionary(Of String, String)
                        log_dict.Add("ai_id", request_id.ToString)
                        log_dict.Add("request_id", request_id.ToString)
                        log_dict.Add("admin_id", su.getId)
                        log_dict.Add("owner_id", http_req_form_owner)
                        log_dict.Add("prev_value", ("Due:" & duedb.ToString("dd/MMM/yyyy") & " Owner:" & currentOwner & " Status:" & currentStatus & " Desc:" & currentDesc).Replace("'", ""))
                        log_dict.Add("new_value", changes.Replace("'", ""))
                        log_dict.Add("event", "AI_UPDATED")
                        log_dict.Add("detail", "Some detail here...")

                        newLog.LogWrite(log_dict)

                    End If

                End If

                Dim redirectTo As String
                redirectTo = syscfg.getSystemUrl + "sap_req.aspx?id=" + redirect_req.ToString
                Response.Redirect(redirectTo, False)
                'response.write(changes)
            End If

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

            'ROWS ITERATION
            'sql_req = "SELECT * FROM requests WHERE id=" + request_id.ToString
            sql_ais = "SELECT * FROM actionitems WHERE id=" + request_id.ToString + extra

            'dbcomm_req = New OleDbCommand(sql_req, dbconn)
            dbcomm_ais = New OleDbCommand(sql_ais, dbconn)

            'ROWS ITERATION
            '#####TODO:#CHECK#IF#REQUEST#EXIST######
            dbread_ais = dbcomm_ais.ExecuteReader()
            If dbread_ais.HasRows Then

                dbread_ais.Read()

                ai_id.Text = ai_id.Text + Convert.ToInt64(dbread_ais.GetValue(0)).ToString
                'ainumber.Text = req_id.Text

                'req_detail.Text = dbread_req.GetString(4)
                'subj.InnerText = req_detail.Text

                'req_description.Text = dbread_ais.GetString(2)

                'If Len(req_description.Text) > 250 Then
                '    req_description.Text = Mid(req_description.Text, 1, 250) & "..."
                'End If

                'Dim mailHTML As String = dbread_req.GetString(3)
                'mail_detail.InnerHtml = HtmlDecode(mailHTML.Replace("&#34;", """").Replace("&#39;", "'"))

                owner.Value = dbread_ais.GetString(6)

                status.Value = dbread_ais.GetString(5)

                descr.InnerText = dbread_ais.GetString(2).Replace("&#39;", "'")

                Dim req_created As Date
                Dim req_duedate As Date
                req_created = dbread_ais.GetDateTime(3)
                If dbread_ais.IsDBNull(4) Then
                    duedate.Value = ""
                Else
                    req_duedate = dbread_ais.GetDateTime(4)
                    duedate.Value = req_duedate.ToString("dd/MMM/yyyy")
                End If

                link_req_id.HRef = syscfg.getSystemUrl + "sap_req.aspx?id=" + dbread_ais.GetInt64(1).ToString
				
				link_del_ai.HRef = syscfg.getSystemUrl + "sap_ai_del.aspx?id=" + Convert.ToInt64(dbread_ais.GetValue(0)).ToString + "&req=" + dbread_ais.GetInt64(1).ToString
				
				dbcomm_req = New OleDbCommand("SELECT * FROM requests WHERE id=" + dbread_ais.GetInt64(1).ToString, dbconn)
				dbread_req = dbcomm_req.ExecuteReader()
				If dbread_req.HasRows Then
					dbread_req.Read()
					
					If dbread_req.IsDBNull(6) Then
						max_date.Text = "false"
					Else
						req_duedate = dbread_req.GetDateTime(6)
						max_date.Text = "'" + req_duedate.ToString("yyyy/MM/dd") + "'"
						request_due.Text = "Request Due: " + req_duedate.ToString("dd/MMM/yyyy")
					End If
				End If

            End If

            dbread_ais.Close()
            dbconn.Close()

        End If

    End Sub

End Class
