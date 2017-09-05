Imports System.Data.OleDb
Imports commonLib

Partial Class sap_test
    Inherits System.Web.UI.Page

    Protected Function sum_filter(filter As String) As Integer
        Dim sysConfiguration As New SystemConfiguration
        Dim actions As New SapActions
        Dim dbconn As OleDbConnection
        Dim dbcomm As OleDbCommand
        Dim dbread As OleDbDataReader

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        Dim extra_where As String = "(status = 'CR' OR status = 'ND') AND ai_count = 0"
        Select Case filter
            Case "ur"
                extra_where = "(status = 'CR' OR status = 'ND') AND ai_count = 0"
            Case "nd"
                extra_where = "status = 'ND'"
            Case "du"
                extra_where = "(status = 'CR' OR status='PD') AND ai_count > 1"
            Case "ap"
                extra_where = "status = 'CR'"
            Case "ex"
                extra_where = "status = 'EX'"
            Case "dl"
                extra_where = "status = 'DL'"
            Case "od"
                extra_where = "status = 'IP'"
        End Select

        Dim sql As String = "SELECT *, (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id) AS ai_count FROM requests WHERE " & extra_where & " ORDER BY id DESC"
        Dim res As Integer = 0
        dbcomm = New OleDbCommand(sql, dbconn)
        dbread = dbcomm.ExecuteReader()

        If dbread.HasRows Then

            While dbread.Read()
                Dim current_req_id As Long = Convert.ToInt64(dbread.GetValue(0))
                Dim req_duedate As Date = actions.requestGetAIsLastDueDate(current_req_id)
                Dim req_missing_days As Integer = 0
                If req_duedate = Nothing Then
                    req_missing_days = -1
                    req_duedate = Today()
                Else
                    req_missing_days = DateDiff(DateInterval.Day, Today.Date, req_duedate.Date)
                End If

                If filter <> "od" Or (filter = "od" And req_missing_days < 7 And req_missing_days > -1) Then
                    res = res + 1
                End If

            End While

        End If

        Return res
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim actions As New SapActions
        'Dim anal As New SapAnalytics

        Response.Write(sum_filter("ur") & " Unassigned Request <a href='http://rtm-bmo.bue.sap.corp:8888/sap_main.aspx?f=ur'>View Items</a><br>")
        Response.Write(sum_filter("nd") & " Need Data <a href='http://rtm-bmo.bue.sap.corp:8888/sap_main.aspx?f=nd'>View Items</a><br>")
        Response.Write(sum_filter("ap") & " Accept Pending <a href='http://rtm-bmo.bue.sap.corp:8888/sap_main.aspx?f=ap'>View Items</a><br>")
        Response.Write(sum_filter("ex") & " Extending Pending for Appr/Rej <a href='http://rtm-bmo.bue.sap.corp:8888/sap_main.aspx?f=ex'>View Items</a><br>")
        Response.Write(sum_filter("dl") & " Deliver Pending for Appr/Rej <a href='http:rtm-bmo.bue.sap.corp:8888/sap_main.aspx?f=dl'>View Items</a><br>")
        Response.Write(sum_filter("od") & " Overdue within a week <a href='http://rtm-bmo.bue.sap.corp:8888/sap_main.aspx?f=od'>View Items</a><br>")


        Response.End()
    End Sub

End Class
