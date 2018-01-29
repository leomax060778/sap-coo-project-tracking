Imports System.Data.OleDb
Imports commonLib

Public Class SapActions

    Dim utilCommon As New commonLib.Utils
    Dim sysConfiguration As New SystemConfiguration
    Dim userCommon As New commonLib.SapUser

    Public Function getRequestorIdFromRequestId(ByVal req_id As Integer) As String
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String
        Dim result As String = Nothing
        Dim user As New SapUser

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        'CHECK AI ACTUAL STATUS
        sql_req = "SELECT * FROM requests WHERE id=" + req_id.ToString
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()

        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.GetString(1)
        End If

        dbconn.Close()

        Return result

    End Function

    Public Function requestIsUnset(ByVal id As String) As Boolean
        Dim result As Boolean = False
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        sql_req = "SELECT due FROM requests WHERE id=" & id

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.IsDBNull(0)
        End If

        dbconn.Close()

        Return result

    End Function

    Public Sub requestSetStatus(ByVal id As String, ByVal initial As String, ByVal final As String)
        Dim result As Boolean = False
        Dim dbconn As OleDbConnection
        Dim dbcomm_req, dbcomm_req_update As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req, sql_update As String

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        sql_req = "SELECT * FROM requests WHERE id=" & id '& " AND status='" & initial & "'"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()

        If dbread_req.HasRows Then
            'dbread_req.Read()
            'result = dbread_req.IsDBNull(0)
            sql_update = "UPDATE requests SET status='" & final & "' WHERE id=" & id
            dbcomm_req_update = New OleDbCommand(sql_update, dbconn)
            dbcomm_req_update.ExecuteNonQuery()
        End If

        dbread_req.Close()
        dbconn.Close()
    End Sub

    Public Sub completeRequest(ByVal id As String)
        Dim result As Boolean = False
        Dim dbconn As OleDbConnection
        Dim dbcomm_req, dbcomm_req_update As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req, sql_update As String

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()
        sql_req = "SELECT * FROM actionitems WHERE request_id=" & id & " AND (status <> 'XX' AND status <> 'CP')"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()

        'IF ALL AI ARE DELETED OR COMPLETED
        If Not dbread_req.HasRows Then
            'dbread_req.Read()
            'result = dbread_req.IsDBNull(0)
            sql_update = "UPDATE requests SET status='CP' WHERE id=" & id
            dbcomm_req_update = New OleDbCommand(sql_update, dbconn)
            dbcomm_req_update.ExecuteNonQuery()
        End If

        dbread_req.Close()
        dbconn.Close()
    End Sub

    'GETS THE LAST DUE DATE FOR THIS REQUEST
    Public Function requestGetAIsLastDueDate(ByVal id As String) As Date
        Dim result As Date = Nothing
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        'GET THE DUE FROM THE LAST AI
        sql_req = "SELECT due FROM actionitems WHERE request_id=" & id.ToString & " ORDER BY due DESC;"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()

        If dbread_req.HasRows Then
            dbread_req.Read()
            If Not dbread_req.IsDBNull(0) Then
                result = dbread_req.GetDateTime(0)
            End If
        End If

        'GET THE DUE FROM THE REQUEST
        sql_req = "SELECT due FROM requests WHERE id=" & id.ToString & ";"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()

        If dbread_req.HasRows Then
            dbread_req.Read()
            If Not dbread_req.IsDBNull(0) Then
                If result = Nothing Or result < dbread_req.GetDateTime(0) Then
                    result = dbread_req.GetDateTime(0)
                End If
            End If
        End If

        dbconn.Close()

        Return result

    End Function

    Public Sub deliverAI(ByVal ai_id As String, ByVal descr As String, ByVal fileName1 As String, ByVal fileName2 As String, ByVal fileName3 As String, ByVal fileName4 As String, ByVal fileName5 As String)
        Dim users As New SapUser
        Dim link As New Linker
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_ais As OleDbDataReader
        Dim sql_req As String
        Dim deliveryDate As String
        Dim fn As String = ""

        If fileName1 <> "" Then
            fn = fn & Chr(34) & fileName1 & Chr(34)
        End If
        If fileName2 <> "" Then
            fn = fn & "," & Chr(34) & fileName2 & Chr(34)
        End If
        If fileName3 <> "" Then
            fn = fn & "," & Chr(34) & fileName3 & Chr(34)
        End If
        If fileName4 <> "" Then
            fn = fn & "," & Chr(34) & fileName4 & Chr(34)
        End If
        If fileName5 <> "" Then
            fn = fn & "," & Chr(34) & fileName5 & Chr(34)
        End If


        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        deliveryDate = Today.Year.ToString + "-" + Today.Month.ToString + "-" + Today.Day.ToString

        sql_req = "UPDATE actionitems SET delivery_date='" + deliveryDate + "', delivery_file='" + fn + "', delivery_text='" + descr + "', status='DL' WHERE id=" + ai_id
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbcomm_req.ExecuteNonQuery()

        sql_req = "SELECT * FROM actionitems WHERE id=" + ai_id
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_ais = dbcomm_req.ExecuteReader()

        If dbread_ais.HasRows Then
            dbread_ais.Read()

            Dim newMail As New MailTemplate
            Dim mail_dict As New Dictionary(Of String, String)

            Dim request As Integer = dbread_ais.GetInt64(1)
            Dim requestor As String = getRequestorIdFromRequestId(request)
            Dim owner As String = dbread_ais.GetString(6)

            mail_dict.Add("mail", "CP") 'NEW AI CREATED
            mail_dict.Add("to", userCommon.getMailById(requestor))
            mail_dict.Add("{ai_id}", ai_id)
            mail_dict.Add("{owner}", userCommon.getNameById(owner) & "(" & owner & ")")
            mail_dict.Add("{description}", dbread_ais.GetString(2))
            mail_dict.Add("{detail}", dbread_ais.GetString(2))
            mail_dict.Add("{duedate}", dbread_ais.GetDateTime(4).ToString("dd/MMM/yyyy"))
            mail_dict.Add("{delivery}", Now.Date.ToString("dd/MMM/yyyy"))
            mail_dict.Add("{requestor_name}", userCommon.getNameById(requestor))
            mail_dict.Add("{app_link}", sysConfiguration.getSystemUrl)
            mail_dict.Add("{contact_mail_link}", "mailto:" & userCommon.getAdminMail & "?subject=Questions about the report")
            mail_dict.Add("{subject}", "Delivery Notice | AI#" & ai_id)

            If fileName1 <> "" Then
                mail_dict.Add("{delivery_link1}", sysConfiguration.getSystemUrl + "delivery.ashx?file=" + utilCommon.encode(fileName1))
                mail_dict.Add("{d1}", "block")
                mail_dict.Add("{filename1}", fileName1)
            Else
                mail_dict.Add("{d1}", "none")
            End If
            If fileName2 <> "" Then
                mail_dict.Add("{delivery_link2}", sysConfiguration.getSystemUrl + "delivery.ashx?file=" + utilCommon.encode(fileName2))
                mail_dict.Add("{d2}", "block")
                mail_dict.Add("{filename2}", fileName2)
            Else
                mail_dict.Add("{d2}", "none")
            End If
            If fileName3 <> "" Then
                mail_dict.Add("{delivery_link3}", sysConfiguration.getSystemUrl + "delivery.ashx?file=" + utilCommon.encode(fileName3))
                mail_dict.Add("{d3}", "block")
                mail_dict.Add("{filename3}", fileName3)
            Else
                mail_dict.Add("{d3}", "none")
            End If
            If fileName4 <> "" Then
                mail_dict.Add("{delivery_link4}", sysConfiguration.getSystemUrl + "delivery.ashx?file=" + utilCommon.encode(fileName4))
                mail_dict.Add("{d4}", "block")
                mail_dict.Add("{filename4}", fileName4)
            Else
                mail_dict.Add("{d4}", "none")
            End If
            If fileName5 <> "" Then
                mail_dict.Add("{d5}", "block")
                mail_dict.Add("{delivery_link5}", sysConfiguration.getSystemUrl + "delivery.ashx?file=" + utilCommon.encode(fileName5))
                mail_dict.Add("{filename5}", fileName5)
            Else
                mail_dict.Add("{d5}", "none")
            End If

            mail_dict.Add("{accept_link}", sysConfiguration.getSystemUrl + "sap_completed.aspx?id=" + link.enLink(ai_id.ToString))
            mail_dict.Add("{uncompleted_link}", sysConfiguration.getSystemUrl + "sap_uncompleted.aspx?id=" + link.enLink(ai_id.ToString))

            newMail.SendNotificationMail(mail_dict)

            '////////////////////////////////////////////////////////////
            'INSERT LOG HERE
            '////////////////////////////////////////////////////////////
            'ai_id, request_id, owner_id, requestor_id, admin_id, event, prev_value, new_value, detail

            'EVENT: AI_CREATED [R1]

            Dim newLog As New commonLib.Logging

            Dim log_dict As New Dictionary(Of String, String)
            log_dict.Add("ai_id", ai_id)
            log_dict.Add("request_id", request.ToString)
            log_dict.Add("admin_id", userCommon.getId)
            log_dict.Add("owner_id", owner)
            log_dict.Add("requestor_id", requestor)
            log_dict.Add("new_value", "")
            log_dict.Add("event", "AI_DELIVERED")
            log_dict.Add("detail", fn)

            newLog.LogWrite(log_dict)

            'Update a Lumira AI
            Dim lumiraReport As New LumiraReports
            Dim lumira_ai As New Dictionary(Of String, String)

            lumira_ai.Add("ai_id", ai_id)
            lumira_ai.Add("req_id", request)
            lumira_ai.Add("delivered", Date.Now.ToString("yyyy/MM/dd HH:mm:ss"))

            lumiraReport.LogActionItemReport(ai_id, lumira_ai)
            'End update Lumira AI

        End If

        dbconn.Close()

    End Sub

End Class
