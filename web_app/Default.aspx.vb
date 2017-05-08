Imports System.Data.OleDb
Imports System
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports SysConfig

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim syscfg As New SysConfig
        Dim redirectTo As String
        redirectTo = syscfg.getSystemUrl + "sap_main.aspx"
        Response.Redirect(redirectTo, False)

        ''Dim dbconn, sql, dbcomm, dbread
        ''dbconn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; data source=" & Server.MapPath("northwind.mdb"))
        ''dbconn.Open()
        ''sql = "SELECT * FROM customers"
        ''dbcomm = New OleDbCommand(sql, dbconn)
        ''dbread = dbcomm.ExecuteReader()
        ''customers.DataSource = dbread
        ''customers.DataBind()
        ''dbread.Close()
        ''dbconn.Close()

        'Dim dbconn As OleDbConnection
        'Dim dbcomm_req, dbcomm_ais As OleDbCommand
        'Dim dbread_req, dbread_ais As OleDbDataReader
        'Dim sql_req, sql_ais As String

        'dbconn = New OleDbConnection("Provider=SAOLEDB;Database=saptareas;Uid=root;Pwd=root;")
        'dbconn.Open()

        ''############INSERT#NEW#ROW#############
        ''sql = "INSERT INTO requests (requestor, mail, subject, detail, due) VALUES ('pablo@sap.com', 'Dear, Juan Perez. Need a new slaptron for next tuesday. Thanks!', 'New Slaptron', 'Next tuesday', '2015-02-25')"

        ''############ROWS#ITERATION#############
        'sql_req = "SELECT * FROM requests WHERE id=2"
        'sql_ais = "SELECT * FROM actionitems WHERE request_id=2"

        'dbcomm_req = New OleDbCommand(sql_req, dbconn)
        'dbcomm_ais = New OleDbCommand(sql_ais, dbconn)

        ''############INSERT#NEW#ROW#############
        ''dbcomm.ExecuteScalar()

        ''############ROWS#ITERATION#############
        'dbread_req = dbcomm_req.ExecuteReader()
        'If dbread_req.HasRows Then

        '    dbread_req.Read()

        '    req_id.Text = Convert.ToInt64(dbread_req.GetValue(0))
        '    req_detail.Text = dbread_req.GetString(4)
        '    req_description.Text = dbread_req.GetString(5)

        '    Dim req_created As Date
        '    Dim req_duedate As Date
        '    req_created = dbread_req.GetDateTime(2)
        '    req_duedate = dbread_req.GetDateTime(6)

        '    req_created_year.Text = req_created.Year
        '    req_created_month.Text = MonthName(req_created.Month, True)
        '    req_created_day.Text = req_created.Day

        '    req_duedate_year.Text = req_duedate.Year
        '    req_duedate_month.Text = MonthName(req_duedate.Month, True)
        '    req_duedate_day.Text = req_duedate.Day

        '    Dim tabla As HtmlTable
        '    tabla = FindControl("ai_list")
        '    'Dim tabla As Table = Page.FindControl("table#ai-list")


        '    dbread_ais = dbcomm_ais.ExecuteReader()
        '    If dbread_ais.HasRows Then
        '        While dbread_ais.Read()

        '            Dim tRow As New HtmlTableRow

        '            '<td>151</td>
        '            Dim tCell_id As New HtmlTableCell
        '            tCell_id.InnerText = Convert.ToInt64(dbread_ais.GetValue(0))
        '            tRow.Cells.Add(tCell_id)

        '            '<td class="ai-desc">We need a snack machine to be installed in the new office. I'll be back by 14/02/2015 and I need it ASAP. Please make sure that it has a variety of healthy snacks too. Thanks!</td>
        '            Dim tCell_des As New HtmlTableCell
        '            tCell_des.InnerText = dbread_ais.GetString(2)
        '            tRow.Cells.Add(tCell_des)

        '            '<td>Owner</td>
        '            Dim tCell_own As New HtmlTableCell
        '            tCell_own.InnerText = dbread_ais.GetString(6)
        '            tRow.Cells.Add(tCell_own)

        '            '<td>Jan 25, 2015</td>
        '            Dim tCell_ctd As New HtmlTableCell
        '            Dim ais_created As Date
        '            ais_created = dbread_ais.GetDateTime(3)
        '            tCell_ctd.InnerText = ais_created.Date.ToString
        '            tRow.Cells.Add(tCell_ctd)

        '            '<td>Feb 14, 2015</td>
        '            Dim tCell_due As New HtmlTableCell
        '            Dim ais_duedate As Date
        '            'ais_duedate = dbread_ais.GetDateTime(4)
        '            tCell_due.InnerText = "NULL" 'ais_duedate.Date.ToString
        '            tRow.Cells.Add(tCell_due)

        '            '<td>Status</td>
        '            Dim tCell_sts As New HtmlTableCell
        '            tCell_sts.InnerText = dbread_ais.GetString(5)
        '            tRow.Cells.Add(tCell_sts)

        '            '<td><a href="#">Create</a></td>
        '            Dim tCell_btn As New HtmlTableCell
        '            tCell_btn.InnerText = "Edit"
        '            tRow.Cells.Add(tCell_btn)

        '            tabla.Rows.Add(tRow)

        '        End While
        '    End If
        '    dbread_ais.Close()
        'End If

        'dbread_req.Close()
        'dbconn.Close()
    End Sub

End Class
