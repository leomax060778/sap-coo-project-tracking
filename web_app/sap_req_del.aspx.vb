Imports System.Data.OleDb
Imports commonLib
Imports MailTemplate
Imports SysConfig

Partial Class sap_req_del
    Inherits System.Web.UI.Page

    Dim sysConfiguration As New SystemConfiguration
    Dim userCommon As New commonLib.SapUser

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim actions As New SapActions
        Dim ro As String = userCommon.getRole()

        Dim http_req_id As String = Request("id")
        Dim request_id As Integer

        If Not String.IsNullOrEmpty(http_req_id) And Integer.TryParse(http_req_id, request_id) And ro <> "OW" Then

            Dim dbconn As OleDbConnection
            Dim dbcomm_req As OleDbCommand
            Dim sql_req As String

            '#####TODO:#CHECK#IF#DB#EXIST###########
            dbconn = New OleDbConnection(sysConfiguration.getConnection)
            dbconn.Open()

            'LOG INFORMATION
            Dim log_rq_id As String = http_req_id
            Dim log_detail As String = "Some detail here..."
            Dim log_owner As String = "Current USER here..."
            Dim log_record As Boolean = False

            'CHANGE REQUEST STATUS TO CANCELED
            sql_req = "UPDATE requests SET status='XX' WHERE id=" + http_req_id
            dbcomm_req = New OleDbCommand(sql_req, dbconn)
            dbcomm_req.ExecuteNonQuery()

            'CHANGE AIS STATUS TO CANCELED
            sql_req = "UPDATE actionitems SET status='XX' WHERE request_id=" + http_req_id
            dbcomm_req = New OleDbCommand(sql_req, dbconn)
            dbcomm_req.ExecuteNonQuery()

            '###################################################
            'WOULD HAVE TO SEND EMAIL TO EACH OWNER TO INFORMING
            'WOULD HAVE TO LOG CHANGE STATUS FROM ?? TO XX
            '###################################################

            Dim newLog As New Logging

            Dim log_dict As New Dictionary(Of String, String)
            log_dict.Add("request_id", http_req_id)
            log_dict.Add("admin_id", userCommon.getId)
            log_dict.Add("event", "RQ_CANCELED")
            log_dict.Add("detail", "Request canceled")

            newLog.LogWrite(log_dict)

            Response.Redirect(sysConfiguration.getSystemUrl + "sap_main.aspx", False)

            dbconn.Close()

        Else
            Response.Redirect(sysConfiguration.getSystemUrl + "sap_main.aspx", False)
            'REDIRECT ME OUT OF HERE
            '#####TODO#HANDLE#EXCEPTION#########
            'Throw New Exception("Request ID not found!")
        End If

    End Sub
End Class
