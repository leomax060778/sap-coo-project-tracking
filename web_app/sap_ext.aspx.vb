Imports System.Data.OleDb
Imports commonLib
Imports System.Globalization

Partial Class Default2
    Inherits System.Web.UI.Page

    Dim sysConfiguration As New SystemConfiguration
    Dim userCommon As New commonLib.SapUser
    Dim utilCommon As New commonLib.Utils

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'CHECK IF DB EXISTS
        '#####TODO###################

        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
        Dim dbread_req, dbread_ais As OleDbDataReader
        Dim sql, sql_ais As String

        Dim users As New SapUser
        Dim actions As New SapActions
        Dim utils As New Utils

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        'CLEAR FORM
        reason.InnerText = ""
        duedate.Value = ""

        'REQUEST FORM
        'If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        If Not IsPostBack Then

            '#####TODO:#CHECK#VALUES################
            Dim http_req_form_ai_id As String = Request.Form("form_ai_id")
            Dim http_req_form_reason As String = Request.Form("reason")
            Dim http_req_form_duedate As String = Request.Form("duedate")

            Dim due_parse_success As Boolean
            Dim ai_new_due As Date
            Dim provider As CultureInfo = CultureInfo.InvariantCulture
            Dim style As DateTimeStyles = DateTimeStyles.AssumeUniversal
            due_parse_success = Date.TryParse(http_req_form_duedate, ai_new_due)
            'ai_new_due = ai_new_due.Date

            'CHECK AI ACTUAL STATUS
            sql_ais = "SELECT * FROM actionitems WHERE id=" + http_req_form_ai_id + " AND owner='" + userCommon.getId + "'"
            dbcomm_ais = New OleDbCommand(sql_ais, dbconn)
            dbread_ais = dbcomm_ais.ExecuteReader()

            If dbread_ais.HasRows And due_parse_success Then

                dbread_ais.Read()
                Dim ai_created As Date = dbread_ais.GetDateTime(3)
                Dim ai_owner As String = dbread_ais.GetString(6)
                Dim request_id As Integer = dbread_ais.GetInt64(1)
                Dim requestor_id As String = actions.getRequestorIdFromRequestId(request_id)
                Dim ai_descr As String = dbread_ais.GetString(2)
                Dim ai_old_due As Date
                If Not dbread_ais.IsDBNull(4) Then
                    ai_old_due = dbread_ais.GetDateTime(4)
                Else
                    ai_old_due = Today.Date
                End If
                Dim ai_old_status As String = dbread_ais.GetString(5)
                Dim ai_missing_days As Integer = DateDiff(DateInterval.Day, Today.Date, ai_old_due.Date)

                'IF AI STATUS VERIFY EXTENSION CONDITION THEN SETS STATUS TO NE
                If ai_old_status <> "CF" And ai_old_status <> "DL" And ai_old_status <> "NE" And ai_missing_days > 0 And ai_old_due < ai_new_due Then
                    'DUEDATE EXTENSION
                    'sql = "UPDATE actionitems SET extension='" + http_req_form_duedate + "', ext_desc='" + http_req_form_reason + "', status='NE' WHERE id =" + http_req_form_ai_id
                    Dim extensionDate As String = ai_new_due.Year.ToString + "-" + ai_new_due.Month.ToString + "-" + ai_new_due.Day.ToString

                    sql = "UPDATE actionitems SET extension='" + extensionDate + "', ext_desc='" + http_req_form_reason + "', status='NE' WHERE id =" + http_req_form_ai_id
                    dbcomm = New OleDbCommand(sql, dbconn)
                    dbcomm.ExecuteScalar()

                    '////////////////////////////////////////////////////////////
                    'CREATE EMAIL AND SEND TO REQUESTOR
                    '////////////////////////////////////////////////////////////
                    Dim newMail As New MailTemplate

                    'NEED ENCRYPTION HERE
                    Dim link As New Linker

                    Dim mail_dict As New Dictionary(Of String, String)
                    mail_dict.Add("mail", "NE") 'AI EXTENSION
                    mail_dict.Add("to", userCommon.getMailById(requestor_id))
                    mail_dict.Add("{ai_id}", http_req_form_ai_id)
                    mail_dict.Add("{owner}", userCommon.getNameById(ai_owner) & "(" & ai_owner & ")")
                    mail_dict.Add("{description}", ai_descr) 'MAIL SUBJECT / AI DESCRIPTION
                    mail_dict.Add("{duedate}", utils.formatDateToSTring(ai_old_due))
                    mail_dict.Add("{extension}", http_req_form_duedate)
                    mail_dict.Add("{reason}", http_req_form_reason)
                    mail_dict.Add("{accept_link}", sysConfiguration.getSystemUrl + "sap_accept_due.aspx?id=" + link.enLink(http_req_form_ai_id))
                    mail_dict.Add("{reject_link}", sysConfiguration.getSystemUrl + "sap_reject_due.aspx?id=" + link.enLink(http_req_form_ai_id))
                    mail_dict.Add("{requestor_name}", userCommon.getNameById(requestor_id))
                    mail_dict.Add("{app_link}", sysConfiguration.getSystemUrl)
                    mail_dict.Add("{contact_mail_link}", "mailto:" & userCommon.getAdminMail & "?subject=Questions about the report")
                    mail_dict.Add("{subject}", "Extension Requested for AI#" & http_req_form_ai_id)

                    newMail.SendNotificationMail(mail_dict)

                    '////////////////////////////////////////////////////////////
                    'INSERT LOG HERE
                    '////////////////////////////////////////////////////////////

                    'Update a Lumira AI
                    Dim lumiraReport As New LumiraReports
                    Dim lumira_ai As New Dictionary(Of String, String)

                    lumira_ai.Add("ai_id", http_req_form_ai_id)
                    lumira_ai.Add("req_id", request_id)
                    lumira_ai.Add("extension", Date.Now.ToString("yyyy/MM/dd HH:mm:ss"))

                    'lumira_ai.Add("original_due", ai_old_due.ToString())
                    'Date.ParseExact(ai_old_due, "yyyy/MM/dd", New CultureInfo("en-US")))
                    lumira_ai.Add("due", extensionDate)

                    lumiraReport.LogActionItemReport(http_req_form_ai_id, lumira_ai)
                    'End update Lumira AI

                    'EVENT: AI_EXTENSION [R5]

                    Dim newLog As New Logging

                    Dim log_dict As New Dictionary(Of String, String)
                    log_dict.Add("ai_id", http_req_form_ai_id)
                    log_dict.Add("request_id", request_id.ToString)
                    log_dict.Add("admin_id", userCommon.getId)
                    log_dict.Add("owner_id", ai_owner)
                    log_dict.Add("requestor_id", requestor_id)
                    log_dict.Add("prev_value", ai_old_due)
                    log_dict.Add("new_value", http_req_form_duedate)
                    log_dict.Add("event", "AI_EXTENSION")
                    log_dict.Add("detail", userCommon.getName + ", requests an extension [" + DateDiff("d", ai_created, ai_old_due).ToString + "+" + DateDiff("d", ai_old_due, ai_new_due).ToString + "days] until " + ai_new_due.ToString("dddd, MMMM dd, yyyy"))

                    newLog.LogWrite(log_dict)

                End If

                Dim redirectTo As String
                redirectTo = sysConfiguration.getSystemUrl + "sap_owner.aspx"
                Response.Redirect(redirectTo, False)

            Else

                'AI DOES NOT EXIST

            End If

        End If

        'REQUEST ID
        Dim http_ai_id As String = Request("id")
        Dim ai_id As Integer
        Dim i As Integer

        If String.IsNullOrEmpty(http_ai_id) Or Not Integer.TryParse(http_ai_id, i) Then
            Response.Redirect(".\sap_error.aspx", False)
        Else
            Integer.TryParse(http_ai_id, ai_id)
        End If

        Dim linked As New Linker

        If i > 1000000 Then
            ai_id = linked.deLink(http_ai_id)
        End If

        sql_ais = "SELECT * FROM actionitems WHERE id=" + ai_id.ToString + " AND owner='" + userCommon.getId + "'"

        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)

        label_ai_id.Text = ai_id.ToString
        form_ai_id.Value = ai_id.ToString

        dbread_ais = dbcomm_ais.ExecuteReader()

        If dbread_ais.HasRows Then

            dbread_ais.Read()

            Dim tRow As New HtmlTableRow

            'id
            Dim current_ai_id As Long
            'current_ai_id = Convert.ToInt64(dbread_ais.GetValue(0))
            current_ai_id = dbread_ais.GetInt64(0)
            ai_tbl_id.Text = current_ai_id

            'desc
            ai_tbl_desc.Text = dbread_ais.GetString(2)

            'created
            Dim ai_created As Date
            ai_created = dbread_ais.GetDateTime(3)
            ai_tbl_created.Text = ai_created.Date.ToString("dd/MMM/yyyy")

            'due
            Dim ai_duedate As Date
            If Not dbread_ais.IsDBNull(4) Then
                ai_duedate = dbread_ais.GetDateTime(4)
            Else
                ai_duedate = Today.Date
            End If
            ai_tbl_due.Text = ai_duedate.Date.ToString("dd/MMM/yyyy")

            ai_duedate = ai_duedate.AddDays(1)
            'MIN_DATE MUST BE GREATER THAN ACTUAL DUEDATE
            '#######TODO#################################
            min_date.Text = "'" + ai_duedate.Year.ToString + "/" + ai_duedate.Month.ToString + "/" + ai_duedate.Day.ToString + "'"
            set_date.Text = min_date.Text
            'desc
            ai_tbl_status.Text = utilCommon.ai_Str_Status(dbread_ais.GetString(5))

            dbcomm_req = New OleDbCommand("SELECT * FROM requests WHERE id=" + dbread_ais.GetInt64(1).ToString, dbconn)
            dbread_req = dbcomm_req.ExecuteReader()
            Dim req_duedate As Date
            If dbread_req.HasRows Then
                dbread_req.Read()

                If dbread_req.IsDBNull(6) Then
                    '	max_date.Text = "false"
                Else
                    req_duedate = dbread_req.GetDateTime(6)
                    'max_date.Text = "'" + req_duedate.ToString("yyyy/MM/dd") + "'"
                    'request_due.Text = "Request Due: " + req_duedate.ToString("dd/MMM/yyyy")
                End If
            End If

        Else

            'ERROR: AI DOES NOT EXIST
            dbconn.Close()
            dbread_ais.Close()
            Response.Redirect(sysConfiguration.getSystemUrl + "sap_owner.aspx", False)

        End If

        dbread_ais.Close()
        dbconn.Close()

        'Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
    End Sub

End Class
