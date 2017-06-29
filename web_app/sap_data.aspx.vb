Imports System.Data.OleDb

Partial Class Default2
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'CHECK IF DB EXISTS
        '#####TODO###################
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()


        Dim http_req_id As String = Request("id")
        Dim req_id As Integer
        Dim i As Integer

        If String.IsNullOrEmpty(http_req_id) Or Not Integer.TryParse(http_req_id, i) Then
            Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
        Else
            Integer.TryParse(http_req_id, req_id)
        End If

        'REQUEST FORM
        If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

            '#####TODO:#CHECK#VALUES################
            Dim http_req_form_name As String = Request.Form("check_name")
            Dim http_req_form_descr As String = Request.Form("check_desc")
            Dim http_req_form_duedate As String = Request.Form("check_due")
            Dim http_req_form_detail As String = Request.Form("detail")

            Dim http_req_form_id As String = Request.Form("form_req_id")
            Dim http_req_form_hd_name As String = Request.Form("form_req_name")
            Dim http_req_form_hd_descr As String = Request.Form("form_req_descr")
            'Dim http_req_form_hd_due As String = Request.Form("form_req_duedate")
            Dim http_req_mail As String = Request.Form("requestor_mail")

            'CHANGE REQUEST STATUS TO CREATED
            '#####TODO#######################
            sql_req = "UPDATE requests SET status='ND' WHERE id=" + req_id.ToString
            dbcomm_req = New OleDbCommand(sql_req, dbconn)
            dbcomm_req.ExecuteNonQuery()

            'Update a Lumira Request
            Dim lumiraReport As New LumiraReports
            Dim lumira_request As New Dictionary(Of String, String)

            lumira_request.Add("req_id", req_id)
            lumira_request.Add("need_data", 1)

            lumiraReport.LogRequestReport(req_id, lumira_request)
            'End update Lumira Request

            '////////////////////////////////////////////////////////////
            'CREATE EMAIL AND SEND TO REQUESTOR
            '////////////////////////////////////////////////////////////

            Dim urlBody As String = ""

            If http_req_form_name = "#FFFC6E" Then
                urlBody = "Request%20Name%3A%0D%0A%0D%0A"
            End If
            If http_req_form_descr = "#FFFC6E" Then
                urlBody = urlBody & "Description%3A%0D%0A%0D%0A"
            End If
            If http_req_form_duedate = "#FFFC6E" Then
                urlBody = urlBody & "Due%20Date%3A%0D%0A%0D%0A"
            End If

            'MAIL: "SAP Email A - More info.html"
            Dim newMail As New MailTemplate

            Dim mail_dict As New Dictionary(Of String, String)
            mail_dict.Add("mail", "ND") 'REQ MORE INFORMATION NEEDED
            mail_dict.Add("to", http_req_mail)
            mail_dict.Add("{req_id}", http_req_form_id)
            mail_dict.Add("{requestor_name}", users.getNameByMail(http_req_mail))
            mail_dict.Add("{name}", http_req_form_hd_name)
            mail_dict.Add("{hl_name}", http_req_form_name)
            mail_dict.Add("{description}", http_req_form_hd_descr) 'MAIL SUBJECT / AI DESCRIPTION
            mail_dict.Add("{hl_descr}", http_req_form_descr)
            mail_dict.Add("{duedate}", "Unset")
            mail_dict.Add("{hl_due}", http_req_form_duedate)
            mail_dict.Add("{detail}", http_req_form_detail)
            'for each admin in users.admins
            mail_dict.Add("{reply_mail_link}", "mailto:" & users.getAdminMail & "?subject=RQ#" & http_req_form_id.Trim() & "%20-%20To%20process%20your%20request%20additional%20info%20is%20needed&body=Dear%20Admin%2C%0D%0A%0D%0AHere%20is%20the%20info%20required%3A%0D%0A%0D%0A" & urlBody)
            'next

            newMail.SendNotificationMail(mail_dict)


            '////////////////////////////////////////////////////////////
            'INSERT LOG HERE
            '////////////////////////////////////////////////////////////

            'LOG INFORMATION
            Dim newLog As New LogSAPTareas
            Dim log_dict As New Dictionary(Of String, String)
            log_dict.Add("req_id", http_req_form_id)
            log_dict.Add("admin_id", users.getId)
            log_dict.Add("requestor_id", users.getIdByMail(http_req_mail))
            log_dict.Add("event", "MISSING_INFO")
            log_dict.Add("detail", http_req_form_detail) '.Substring(0, 256)) 'DETAIL FIELD IS CHAR(256)
            'newLog.LogWrite(log_dict)

            Dim redirectTo As String
            redirectTo = ".\sap_main.aspx"
            Response.Redirect(redirectTo, False)
        End If

        sql_req = "SELECT * FROM requests WHERE id=" + req_id.ToString
        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        '############ROWS#ITERATION#############
        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then


            dbread_req.Read()

            'name
            req_tbl_name.Text = dbread_req.GetString(4)

            'desc
            req_tbl_desc.Text = dbread_req.GetString(5)
            If Len(req_tbl_desc.Text) > 350 Then
                req_tbl_desc.Text = Mid(req_tbl_desc.Text, 1, 350) & "..."
            End If

            'duedate
            'Dim req_duedate As Date
            'req_duedate = dbread_req.GetDateTime(6)
            'req_tbl_due.Text = req_duedate.Date
            req_tbl_due.Text = "Unset"

            'HIDDEN DATA
            form_req_id.Value = req_id.ToString
            form_req_name.Value = dbread_req.GetString(4)
            form_req_descr.Value = dbread_req.GetString(5)
            'form_req_duedate.Value = req_id.ToString
            requestor_mail.Value = users.getMailById(dbread_req.GetString(1))

        Else

            'ERROR: RQ DOES NOT EXIST
            dbconn.Close()
            dbread_req.Close()
            Response.Redirect(".\sap_main.aspx", False)

        End If

        dbread_req.Close()
        dbconn.Close()
    End Sub

End Class
