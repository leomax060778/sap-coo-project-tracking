Imports System.Data.OleDb
Imports System.IO

Public Class SysConfig

    Public Function getSendMailStatus() As Boolean 'As Dictionary(Of String, String)
        'Dim result As New Dictionary(Of String, String)
        Dim result As Boolean = True
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT * FROM system WHERE id = 1"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.GetBoolean(3)
        End If

        dbconn.Close()

        Return result
    End Function

    Public Function getMailLastCheck() As Date 'As Dictionary(Of String, String)
        'Dim result As New Dictionary(Of String, String)
        Dim result As Date
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT * FROM system WHERE id = 1"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.GetDateTime(2)
        End If

        dbconn.Close()

        Return result
    End Function

    Public Function hasSentReportToday() As Boolean
        Dim result As Boolean
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT * FROM system WHERE id = 1"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.GetDateTime(4).Date = Today.Date
        End If

        dbconn.Close()

        Return result
    End Function

    Public Function hasSentNotificationToday() As Boolean
        Dim result As Boolean
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT * FROM system WHERE id = 1"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.GetDateTime(5).Date = Today.Date
        End If

        dbconn.Close()

        Return result
    End Function

    Public Sub setMailLastCheck(ByVal dt As Date)
        Dim result As Date
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'sql_ais = "UPDATE actionitems SET status='" + ai_new_status + "' WHERE id=" + ai_id
        sql_req = "UPDATE system SET last_mail_date='" + dt.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE id = 1"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            While dbread_req.Read()
                result = dbread_req.GetDateTime(0)
            End While
        End If

        dbconn.Close()
    End Sub

    Public Sub setSentReportToday()
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim sql_req As String
        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()
        sql_req = "UPDATE system SET sent_report='" + Today.ToString("yyyy-MM-dd") + "' WHERE id = 1"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbcomm_req.ExecuteNonQuery()
        dbconn.Close()
    End Sub

    Public Sub setSentNotificationToday()
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim sql_req As String
        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()
        sql_req = "UPDATE system SET sent_notification='" + Today.ToString("yyyy-MM-dd") + "' WHERE id = 1"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbcomm_req.ExecuteNonQuery()
        dbconn.Close()
    End Sub

    Public Function getConnection() As String
        Dim result As String
        'TESTING CONNECTION
        'result = "Provider=SAOLEDB;UID=root;PWD=root;Server=sap-ais-2;DBN=sap-ais-2;ASTART=No;host=localhost:2638"

        'PRODUCTION
        result = "Provider=SAOLEDB;UID=root;PWD=root;Server=sap-ais-2;DBN=sap-ais-2;ASTART=No;host=arbuesql01.phl.sap.corp:2638"

        Return result
    End Function

    Public Function getSystemMail() As String
        'email account for testing
        'Return "support_planningtool@folderit.net"

        'new email account defined for the tool
        Return "sap_marketing_in_action@sap.com"

    End Function

    Public Function getSystemAdminMail() As String
        Dim users As New SapUser
        Return users.getAdminMail
    End Function

    Public Function getSystemUrl() As String
        Dim result As String
        Select Case System.Environment.UserName
            Case "SiS"
                result = "http://localhost:2952/sap-tareas/"
            Case "Vostro"
                result = "http://localhost:4120/sap/"
            Case "PabloVM"
                result = "http://localhost:????/demochoto/"
            Case Else
                'result = "http://localhost:3542/"
                result = "http://rtm-bmo.bue.sap.corp:8888/"
        End Select
        Return result
    End Function

    Public Function req_Str_Status(ByVal s As String) As String
        Dim result As String
        Select Case s
            Case "PD"
                result = "Pending"
            Case "CR"
                result = "Created"
            Case "ND"
                result = "Need Data"
            Case "CP"
                result = "Completed"
            Case Else
                result = "Unset"
        End Select
        Return result
    End Function

End Class
