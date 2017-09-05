Imports System.Data.OleDb
Imports commonLib

Partial Class sap_ai_data
    Inherits System.Web.UI.Page

    Dim sysConfiguration As New SystemConfiguration
    Dim userCommon As New commonLib.SapUser

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'CHECK IF DB EXISTS
        '#####TODO###################
        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_ais, dbcomm_req As OleDbCommand
        Dim dbread_req, dbread_ais As OleDbDataReader
        Dim sql, sql_ais, sql_req As String
        Dim users As New SapUser
        Dim actions As New SapActions
        Dim utils As New Utils

        Dim http_ai_id As String = Request("id")
        Dim req_id, ai_id As Integer
        Dim i As Integer
        Dim linked As New Linker
        Dim ai_duedate As New Date

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        If String.IsNullOrEmpty(http_ai_id) Or Not Integer.TryParse(http_ai_id, i) Then
            Response.Redirect(".\sap_error.aspx", False)
        Else
            Integer.TryParse(http_ai_id, ai_id)
        End If

        If i > 1000000 Then
            ai_id = linked.deLink(http_ai_id)
        End If

        'REQUEST FORM
        If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

            '#####TODO:#CHECK#VALUES################
            Dim http_ai_form_name As String = Request.Form("check_name")
            Dim http_ai_form_descr As String = Request.Form("check_desc")
            Dim http_ai_form_duedate As String = Request.Form("check_due")
            Dim http_ai_form_detail As String = Request.Form("detail")

            Dim http_ai_form_id As String = Request.Form("form_ai_id")
            Dim http_req_form_id As String = Request.Form("form_req_id")
            Dim http_ai_form_hd_name As String = Request.Form("form_ai_name")
            Dim http_ai_form_hd_descr As String = Request.Form("form_ai_descr")
            Dim http_req_mail As String = Request.Form("requestor_mail")
            Dim http_ai_duedate As String = Request.Form("form_ai_duedate")

            'CHECK IF AI EXISTS AND AI ACTUAL STATUS IS NEED_EXTENSION
            sql_ais = "SELECT * FROM actionitems WHERE id=" + ai_id.ToString '+ " AND extension IS NOT NULL"
            dbcomm_ais = New OleDbCommand(sql_ais, dbconn)
            dbread_ais = dbcomm_ais.ExecuteReader()

            If dbread_ais.HasRows Then

                dbread_ais.Read()

                Dim ai_owner As String = dbread_ais.GetString(6)

                'DUEDATE ACCEPT
                sql = "UPDATE actionitems SET status='ND' WHERE id=" + ai_id.ToString
                dbcomm = New OleDbCommand(sql, dbconn)
                dbcomm.ExecuteScalar()

                Dim newLog As New Logging
                Dim log_dict As New Dictionary(Of String, String)

                log_dict.Add("ai_id", ai_id.ToString)
                log_dict.Add("request_id", dbread_ais.GetInt64(1).ToString)
                log_dict.Add("admin_id", userCommon.getId)
                log_dict.Add("owner_id", dbread_ais.GetString(6))
                log_dict.Add("requestor_id", actions.getRequestorIdFromRequestId(dbread_ais.GetInt64(1)))
                log_dict.Add("prev_value", "PD")
                log_dict.Add("new_value", "IP")
                log_dict.Add("event", "AI_ACCEPTED")
                log_dict.Add("detail", "New due date accepted")

                newLog.LogWrite(log_dict)

                'MAIL: "SAP Email A - More info.html"
                Dim newMail As New MailTemplate
                Dim mail_dict As New Dictionary(Of String, String)
                Dim urlBody As String = ""

                If http_ai_form_name = "#FFFC6E" Then
                    urlBody = "Request%20Name%3A%0D%0A%0D%0A"
                End If
                If http_ai_form_descr = "#FFFC6E" Then
                    urlBody = urlBody & "Description%3A%0D%0A%0D%0A"
                End If
                If http_ai_form_duedate = "#FFFC6E" Then
                    urlBody = urlBody & "Due%20Date%3A%0D%0A%0D%0A"
                End If

                mail_dict.Add("mail", "ND") 'ACTION ITEM MORE INFORMATION NEEDED
                mail_dict.Add("to", http_req_mail)
                mail_dict.Add("{requestor_name}", userCommon.getNameByMail(http_req_mail))
                mail_dict.Add("{ai_owner}", userCommon.getNameById(ai_owner))
                mail_dict.Add("{req_id}", http_ai_form_id)
                mail_dict.Add("{name}", http_ai_form_hd_name)
                mail_dict.Add("{hl_name}", http_ai_form_name)
                mail_dict.Add("{description}", http_ai_form_hd_descr) 'MAIL SUBJECT / AI DESCRIPTION
                mail_dict.Add("{hl_descr}", http_ai_form_descr)
                mail_dict.Add("{duedate}", utils.formatDateToSTring(http_ai_duedate))
                mail_dict.Add("{hl_due}", http_ai_form_duedate)
                mail_dict.Add("{detail}", http_ai_form_detail)
                mail_dict.Add("{reply_mail_link}", "mailto:" & userCommon.getMailById(ai_owner) & "?subject=AI#" & http_ai_form_id.Trim() & "%20-%20To%20process%20your%20request%20additional%20info%20is%20needed&body=Dear%20Admin%2C%0D%0A%0D%0AHere%20is%20the%20info%20required%3A%0D%0A%0D%0A" & urlBody)
                mail_dict.Add("{app_link}", sysConfiguration.getSystemUrl)
                mail_dict.Add("{subject}", userCommon.getNameById(ai_owner) & " is requesting information for AI#" & http_ai_form_id.Trim())

                newMail.SendNotificationMail(mail_dict)

            End If

            Response.Redirect(sysConfiguration.getSystemUrl + "sap_main.aspx", False)

        End If

        'On GET
        sql_ais = "SELECT * FROM actionitems WHERE id=" + ai_id.ToString
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)
        dbread_ais = dbcomm_ais.ExecuteReader()

        If dbread_ais.HasRows Then

            dbread_ais.Read()

            ai_duedate = dbread_ais.GetDateTime(4)

            If ai_duedate = Nothing Then
                ai_duedate = Today()
            End If

            req_id = Convert.ToInt64(dbread_ais.GetValue(1))

            sql_req = "SELECT * FROM requests WHERE id=" + Convert.ToInt64(dbread_ais.GetValue(1)).ToString
            dbcomm_req = New OleDbCommand(sql_req, dbconn)
            dbread_req = dbcomm_req.ExecuteReader()
            dbread_req.Read()

            If dbread_req.HasRows Then

                If Not dbread_req.IsDBNull(9) Then
                    download_link.HRef = "./downloadmail.ashx?mail=" + dbread_req.GetString(9) + ".eml"
                End If

                req_description.Text = "AI# " + ai_id.ToString() + "-" + dbread_ais.GetString(2).Replace("&#39;", "'")
                'dbread_req.GetString(5).Replace("&#39;", "'")
                If Len(req_description.Text) > 250 Then
                    req_description.Text = Mid(req_description.Text, 1, 250) & "..."
                End If

                Dim req_created As Date
                Dim req_duedate As Date = dbread_req.GetDateTime(6)
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

                req_created = dbread_ais.GetDateTime(3)

                'HIDDEN DATA
                form_req_id.Value = req_id.ToString
                form_ai_name.Value = dbread_ais.GetString(2)
                form_ai_descr.Value = dbread_ais.GetString(2)
                form_ai_id.Value = Convert.ToInt64(dbread_ais.GetValue(0))
                form_ai_duedate.Value = ai_duedate.ToString("dd/MMM/yyyy")
                requestor_mail.Value = userCommon.getMailById(dbread_req.GetString(1))

                dbread_req.Close()

            End If

        End If

        dbread_ais.Close()
        dbconn.Close()

    End Sub

End Class
