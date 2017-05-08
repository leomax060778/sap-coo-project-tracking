Imports Microsoft.VisualBasic
Imports System.Data.OleDb
Imports System.Collections.Generic

Public Class LogSAPTareas

    Public Sub LogWrite(ByRef eventData As Dictionary(Of String, String))

        Dim dbconn As OleDbConnection
        Dim dbcomm As OleDbCommand
        Dim sql As String
        Dim sql_fields As String = ""
        Dim sql_values As String = ""

        Dim syscfg As New SysConfig

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
        Dim newlog As New LogSAPTareas
        log_dict.Add("detail", eventData)
        newLog.LogWrite(log_dict)
    End Sub

End Class
