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
Imports common
Imports commonLib
'Imports MailTemplate

'#############################################################################################################
'#############################################################################################################
'NOT IMPLEMENTED / TO DO
'#############################################################################################################
'#############################################################################################################

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim syscfg As New SysConfig

        Dim su As New SapUser
        Dim ro As String = su.getRole()

        Dim actions As New SapActions
        Dim link As New Linker
        Dim dbconn As OleDbConnection
        Dim dbcomm_req, dbcomm_ais As OleDbCommand
        Dim dbread_req, dbread_ais As OleDbDataReader
        Dim sql_req, sql_ais As String

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

        Dim sql_owners As String
        sql_owners = "SELECT * FROM users WHERE id='" & su.getId() & "'"

        dbcomm_ais = New OleDbCommand(sql_owners, dbconn)
        dbread_ais = dbcomm_ais.ExecuteReader()

        While dbread_ais.Read()
            'owner_option.Text = owner_option.Text & "<option value='" & dbread_ais.GetValue(0) & "'>" & dbread_ais.GetValue(1) & "</option>"
            owner.Items.Add(New ListItem(dbread_ais.GetValue(1), dbread_ais.GetValue(0)))
        End While

        Dim extra As String = ""

        sql_ais = "SELECT * FROM actionitems WHERE id=" + request_id.ToString + extra
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)

        dbread_ais = dbcomm_ais.ExecuteReader()
        If dbread_ais.HasRows Then

            dbread_ais.Read()
            sql_req = "SELECT * FROM requests WHERE id=" + Convert.ToInt64(dbread_ais.GetValue(1)).ToString
            dbcomm_req = New OleDbCommand(sql_req, dbconn)
            dbread_req = dbcomm_req.ExecuteReader()
            dbread_req.Read()

            ai_id.Text = ai_id.Text + Convert.ToInt64(dbread_ais.GetValue(0)).ToString

            If Not dbread_req.IsDBNull(9) Then
                download_link.HRef = "./downloadmail.ashx?mail=" + dbread_req.GetString(9) + ".eml"
            End If


            req_description.Text = dbread_req.GetString(5).Replace("&#39;", "'")
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
            owner.Value = dbread_ais.GetString(6)

            status.Value = dbread_ais.GetString(5)

            descr.InnerText = dbread_ais.GetString(2).Replace("&#39;", "'")

            req_created = dbread_ais.GetDateTime(3)
            If dbread_ais.IsDBNull(4) Then
                duedate.Value = ""
            Else
                req_duedate = dbread_ais.GetDateTime(4)
                duedate.Value = req_duedate.ToString("dd/MMM/yyyy")
            End If

            link_ai_id.HRef = syscfg.getSystemUrl + "sap_main.aspx"

            REM dbcomm_req = New OleDbCommand("SELECT * FROM requests WHERE id=" + dbread_ais.GetInt64(1).ToString, dbconn)
            REM dbread_req = dbcomm_req.ExecuteReader()
            REM If dbread_req.HasRows Then
            REM dbread_req.Read()

            REM If dbread_req.IsDBNull(6) Then
            REM max_date.Text = "false"
            REM Else
            REM req_duedate = dbread_req.GetDateTime(6)
            REM max_date.Text = "'" + req_duedate.ToString("yyyy/MM/dd") + "'"
            REM request_due.Text = "Request Due: " + req_duedate.ToString("dd/MMM/yyyy")
            REM End If
            REM End If

        End If

        dbread_ais.Close()
        dbconn.Close()



    End Sub

End Class
