Imports System.Data.OleDb
Imports System.IO

Public Class Logging

    Dim appSettings As New AppSettings

    Public Sub log(ByVal s As String)
        Dim strFile As String = appSettings.pathLogFile
        Using sw As New StreamWriter(File.Open(strFile, FileMode.OpenOrCreate))
            sw.WriteLine("[" + Now.ToString("yyyy.MM.dd hh:mm:ss") + "]" + s)
            sw.Flush()
        End Using
    End Sub

    Public Sub LogWrite(ByRef eventData As Dictionary(Of String, String))
        Dim dbconn As OleDbConnection
        Dim dbcomm As OleDbCommand
        Dim sql As String
        Dim sql_fields As String = ""
        Dim sql_values As String = ""

        Dim syscfg As New SystemConfiguration

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        For Each kvp As KeyValuePair(Of String, String) In eventData
            sql_fields = sql_fields + ", " + kvp.Key
            If kvp.Key = "ai_id" Or kvp.Key = "request_id" Then
                sql_values = sql_values + ", " + kvp.Value
            Else
                sql_values = sql_values + ", '" + kvp.Value + "'"
            End If
        Next kvp

        sql_fields = sql_fields.Remove(0, 2)
        sql_values = sql_values.Remove(0, 2)
        sql = "INSERT INTO log (" + sql_fields + ") VALUES (" + sql_values + ")"
        dbcomm = New OleDbCommand(sql, dbconn)
        System.Diagnostics.Debug.WriteLine(dbcomm.ExecuteScalar())

        dbconn.Close()

    End Sub

    Public Sub LogDescr(ByVal eventData As String)
        Dim log_dict As New Dictionary(Of String, String)
        Dim newlog As New Logging
        log_dict.Add("detail", eventData)
        newlog.LogWrite(log_dict)
    End Sub

End Class
