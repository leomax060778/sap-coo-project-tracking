Imports System.Data.OleDb

Public Class LumiraReports

    Public Sub LogRequestReport(ByRef req_id As String, ByRef requestData As Dictionary(Of String, String))
        Dim dbconn As OleDbConnection
        Dim dbcomm As OleDbCommand
        Dim sqlQuery As String
        Dim syscfg As New SysConfig
        Dim action As String = ""
        Dim sql_fields As String = ""
        Dim sql_values As String = ""

        If (requestData.Count = 0) Then Return

        'Open a connection
        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        action = "insert"
        If Not IsNothing(req_id) And ExistsRequest(req_id) Then
            action = "update"
        End If

        If action = "insert" Then
            For Each kvp As KeyValuePair(Of String, String) In requestData
                sql_fields = sql_fields + ", " + kvp.Key
                If kvp.Key = "request_id" Then
                    sql_values = sql_values + ", " + kvp.Value
                Else
                    sql_values = sql_values + ", '" + kvp.Value + "'"
                End If
            Next kvp

            sql_fields = sql_fields.Remove(0, 2)
            sql_values = sql_values.Remove(0, 2)
            sqlQuery = "INSERT INTO lumira_requests (" + sql_fields + ") VALUES (" + sql_values + ")"
            dbcomm = New OleDbCommand(sqlQuery, dbconn)
            dbcomm.ExecuteNonQuery()
        End If

        If action = "update" Then
            For Each kvp As KeyValuePair(Of String, String) In requestData
                If kvp.Key <> "request_id" Then
                    If sql_fields.Length <= 0 Then
                        sql_fields = kvp.Key + " = " + kvp.Value
                    Else
                        sql_fields = sql_fields + ", " + kvp.Key + " = '" + kvp.Value + "'"
                    End If
                End If
            Next kvp

            sql_fields = sql_fields.Remove(0, 2)
            sqlQuery = "UPDATE lumira_requests SET " + sql_fields + " WHERE req_id =" + req_id
            dbcomm = New OleDbCommand(sqlQuery, dbconn)
            dbcomm.ExecuteNonQuery()
        End If

        'Close the DB connection
        dbconn.Close()

    End Sub

    Public Sub LogActionItemReport(ByRef ai_id As String, ByRef aiData As Dictionary(Of String, String))
        Dim dbconn As OleDbConnection
        Dim dbcomm As OleDbCommand
        Dim sqlQuery As String
        Dim syscfg As New SysConfig

        Dim action As String = ""
        Dim sql_fields As String = ""
        Dim sql_values As String = ""

        If (aiData.Count = 0) Then Return

        'Open a connection
        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        action = "insert"
        If Not IsNothing(ai_id) And ExistsActionItem(ai_id) Then
            action = "update"
        End If

        If action = "insert" Then
            For Each kvp As KeyValuePair(Of String, String) In aiData
                sql_fields = sql_fields + ", " + kvp.Key
                If kvp.Key = "ai_id" Then
                    sql_values = sql_values + ", " + kvp.Value
                Else
                    sql_values = sql_values + ", '" + kvp.Value + "'"
                End If
            Next kvp

            sql_fields = sql_fields.Remove(0, 2)
            sql_values = sql_values.Remove(0, 2)
            sqlQuery = "INSERT INTO lumira_ais (" + sql_fields + ") VALUES (" + sql_values + ")"
            dbcomm = New OleDbCommand(sqlQuery, dbconn)
            dbcomm.ExecuteNonQuery()
        End If

        If action = "update" Then
            For Each kvp As KeyValuePair(Of String, String) In aiData
                If sql_fields.Length <= 0 Then
                    sql_fields = kvp.Key + " = " + kvp.Value
                Else
                    sql_fields = sql_fields + ", " + kvp.Key + " = '" + kvp.Value + "'"
                End If
            Next kvp

            'sql_fields = sql_fields.Remove(0, 2)
            sqlQuery = "UPDATE lumira_ais SET " + sql_fields + " WHERE ai_id =" + ai_id
            dbcomm = New OleDbCommand(sqlQuery, dbconn)
            dbcomm.ExecuteNonQuery()
        End If

        'Close the DB connection
        dbconn.Close()

    End Sub

    Private Function ExistsAI(ByVal ai_id As String) As Boolean

        Dim syscfg As New SysConfig
        Dim dbconn As OleDbConnection
        Dim dbcomm_ais As OleDbCommand
        Dim dbread_ais As OleDbDataReader
        Dim existsRecord As Boolean = False
        Dim sql_req As String

        If IsNothing(ai_id) Then Return False

        Try
            dbconn = New OleDbConnection(syscfg.getConnection)
            dbconn.Open()

            sql_req = "SELECT * FROM lumira_ais WHERE ai_id=" + ai_id
            dbcomm_ais = New OleDbCommand(sql_req, dbconn)
            dbread_ais = dbcomm_ais.ExecuteReader()

            If dbread_ais.HasRows Then
                existsRecord = True
            End If

            dbconn.Close()
        Catch ex As Exception

        End Try

        Return existsRecord
    End Function

    Private Function ExistsRequest(ByVal req_id As String) As Boolean

        Dim syscfg As New SysConfig
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim existsRecord As Boolean = False
        Dim sql_req As String

        If IsNothing(req_id) Then Return False

        Try
            dbconn = New OleDbConnection(syscfg.getConnection)
            dbconn.Open()

            sql_req = "SELECT * FROM lumira_requests WHERE req_id=" + req_id
            dbcomm_req = New OleDbCommand(sql_req, dbconn)
            dbread_req = dbcomm_req.ExecuteReader()

            If dbread_req.HasRows Then
                existsRecord = True
            End If

            dbconn.Close()
        Catch ex As Exception

        End Try

        Return existsRecord

    End Function

    Private Function ExistsActionItem(ByVal ai_id As String) As Boolean

        Dim syscfg As New SysConfig
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim existsRecord As Boolean = False
        Dim sql_req As String

        Try
            dbconn = New OleDbConnection(syscfg.getConnection)
            dbconn.Open()

            sql_req = "SELECT * FROM lumira_ais WHERE ai_id=" + ai_id
            dbcomm_req = New OleDbCommand(sql_req, dbconn)
            dbread_req = dbcomm_req.ExecuteReader()

            If dbread_req.HasRows Then
                existsRecord = True
            End If

            dbconn.Close()
        Catch ex As Exception

        End Try

        Return existsRecord

    End Function

End Class
