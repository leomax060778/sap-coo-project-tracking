Imports System.Data.OleDb

Public Class SysConfig

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
            '    While dbread_req.Read()
            '        result.Add("mail_id", dbread_req.GetString(1))
            '        result.Add("mail_date", dbread_req.GetDateTime(2).Date.ToString)
            '        result.Add("mail_time", dbread_req.GetDateTime(2).TimeOfDay.ToString)
            '        result.Add("mail_date", dbread_req.GetDateTime(2).ToString)
            '    End While
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

    Public Function getConnection() As String
        Dim result As String
        If System.Environment.UserName.StartsWith("ARBUE") Then
            result = "Provider=SAOLEDB;UID=root;PWD=root;Server=sap-ais-2;DBN=sap-ais-2;ASTART=No;host=arbuesql01.phl.sap.corp:2638;"
            'ConfigurationSettings.AppSettings("myConnectionString")
        Else
            result = "Provider=SAOLEDB;UID=root;PWD=root;Server=sap-ais-2;DBN=sap-ais-2;ASTART=No;host=localhost:2638;"
            'result = ConfigurationManager.ConnectionStrings("myConnectionString").ConnectionString
        End If
        Return result
    End Function

    Public Function getSystemMail() As String
        Return "sap_marketing_in_action@sap.com"
        'Return "lhildt@folderit.net"
    End Function

    Public Function getSystemAdminMail() As String
        'Return "lhildt@folderit.net"

        Return "a.eyzaguirre@sap.com"
    End Function

    Public Function getSystemUrl() As String
        Dim result As String
        Select Case System.Environment.UserName
            Case "VAIO"
                result = "http://localhost:1085/sap-tareas/"
            Case "Vostro"
                result = "http://localhost:1065/web/"
            Case "Administrator"
                result = "http://localhost:64601/"
            Case Else
                'result = "http://localhost:3542/"
                result = "http://rtm-bmo.bue.sap.corp:8888/"
        End Select
        Return result
    End Function

    Public Function req_Str_Status(ByVal s As String) As String
        Dim result As String
        Select Case s
            Case "IP"
                result = "In Progress"
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
