Imports Microsoft.VisualBasic

Imports System.Data.OleDb
Imports System
Imports SysConfig
Imports Linker

Public Class SapActions

    Public Function getRequestorIdFromRequestId(ByVal req_id As Integer) As String
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String
        Dim result As String = Nothing

        Dim syscfg As New SysConfig
        Dim user As New SapUser

        dbconn = New OleDbConnection(syscfg.getConnection)
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

    Public Function getCountFromLog(ByVal elementType As String, ByVal id As Integer, ByVal eventType As String) As Integer
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim sql_req As String
        Dim result As Integer

        Dim syscfg As New SysConfig
        Dim user As New SapUser

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'CHECK AI ACTUAL STATUS
        sql_req = "SELECT COUNT(*) FROM log WHERE event='" & eventType & "' AND " & elementType & "=" + id.ToString
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        result = dbcomm_req.ExecuteScalar()

        dbconn.Close()

        Return result

    End Function

    Public Function getFirstDateFromLog(ByVal elementType As String, ByVal id As Integer, ByVal eventType As String, ByRef dateResult As Date) As Boolean
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String
        Dim result As Boolean = False
        Dim syscfg As New SysConfig
        Dim user As New SapUser

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'CHECK AI ACTUAL STATUS
        sql_req = "SELECT time_stamp FROM log WHERE event='" & eventType & "' AND " & elementType & "=" + id.ToString & " ORDER BY time_stamp"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()

        If dbread_req.HasRows Then
            dbread_req.Read()
            dateResult = dbread_req.GetDateTime(0)
            result = True
        End If

        dbconn.Close()

        Return result

    End Function

    Public Sub createNewRequest(ByVal sourceMail As Dictionary(Of String, String))
        'CHECK IF DB EXISTS
        '#####TODO###################
        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_ais As OleDbCommand
        Dim dbread_ais As OleDbDataReader
        Dim sql, sql_ais As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New LogSAPTareas

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        Dim mailFrom As String = sourceMail("from")
        Dim mailSubject As String = sourceMail("subject")
        Dim mailBody As String = sourceMail("body")
        Dim mailDate As String = sourceMail("date")
        Dim mailID As String = sourceMail("id")
        Dim mailRaw As String = sourceMail("raw")
        Dim requestorID As String = users.getIdByMail(mailFrom)

        'ID LAST REQUEST CREATED
        Dim req_id As Long
        Dim ai_id As Long

        'CREATE REQUEST
        'IF THE MAIL HAS AT LEAST ONE OWNER THEN SET STATUS TO IN PROGRESS
        'ELSE THE RQ STATUS IS PENDING...
        Dim reqStatus As String = "PD"
        If sourceMail.Count > 6 Then
            reqStatus = "IP"
        End If
        sql = "INSERT INTO requests (requestor, mail, subject, detail, status) VALUES ('" + requestorID + "', '" + mailRaw + "', '" + mailSubject + "', '" + mailBody + "', '" + reqStatus + "')"
        dbcomm = New OleDbCommand(sql, dbconn)
        dbcomm.ExecuteNonQuery()
        dbcomm.CommandText = "SELECT @@IDENTITY"
        req_id = dbcomm.ExecuteScalar()

        'CREATE AN AI FOR EACH OWNER AND SEND AN EMAIL
        'EACH OWNER IS APPENDED IN THE DICTIONARY AS OWN1, OWN2, OWN3...
        Dim i As Integer
        Dim owner, ownerID As String
        Dim newMail As New MailTemplate
        Dim mail_dict As New Dictionary(Of String, String)
        Dim log_dict As New Dictionary(Of String, String)
        Dim ownersMailStr As String = ""

        For i = 1 To sourceMail.Count - 6 'THE MAIL HAS 6 FIELDS THE REST ARE THE OWNERS
            owner = "own" + i.ToString
            ownerID = users.getIdByMail(sourceMail(owner))

            'CREATE ACTION ITEM
            sql = "INSERT INTO actionitems (request_id, description, owner, status) VALUES (" + req_id.ToString + ", '" + mailSubject + " / " + mailBody + "', '" + ownerID + "', 'IP')"
            dbcomm = New OleDbCommand(sql, dbconn)
            dbcomm.ExecuteNonQuery()
            dbcomm.CommandText = "SELECT @@IDENTITY"
            ai_id = dbcomm.ExecuteScalar()

            'SEND MAIL TO OWNER WITH AIS DETAILS
            mail_dict.Add("mail", "CR") 'AI CREATED
            mail_dict.Add("to", sourceMail(owner))
            mail_dict.Add("{ai_id}", ai_id.ToString)
            mail_dict.Add("{description}", "[" + mailSubject + "] " + mailBody) 'MAIL SUBJECT / AI DESCRIPTION
            mail_dict.Add("{duedate}", "See description")
            mail_dict.Add("{accept_link}", syscfg.getSystemUrl + "sap_accept_new_due.aspx?id=" + eLink.enLink(ai_id.ToString))
			mail_dict.Add("{reject_link}", syscfg.getSystemUrl + "sap_reject_due.aspx?id=" + eLink.enLink(ai_id.ToString))
            mail_dict.Add("{extension_link}", syscfg.getSystemUrl + "sap_ext.aspx?id=" + eLink.enLink(ai_id.ToString))		
            mail_dict.Add("{ai_owner}", users.getNameByMail(sourceMail(owner)))
            mail_dict.Add("{app_link}", syscfg.getSystemUrl)
            mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
			mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + ai_id.ToString)

            newMail.SendNotificationMail(mail_dict)

            log_dict.Add("request_id", req_id.ToString)
            log_dict.Add("ai_id", ai_id.ToString)
            log_dict.Add("requestor_id", requestorID)
            log_dict.Add("admin_id", "system")
            log_dict.Add("owner_id", ownerID)
            log_dict.Add("event", "AI_CREATED")
            log_dict.Add("detail", "Some detail here...")

            newLog.LogWrite(log_dict)

            mail_dict.Clear()
            log_dict.Clear()

            ownersMailStr = ownersMailStr + users.getNameByMail(sourceMail(owner)) + "<br>"
        Next

        '////////////////////////////////////////////////////////////
        ' SEND TO ADMIN WITH REQUEST AND OWNERS AIS DETAILS
        '////////////////////////////////////////////////////////////

        'Dim mail_dict As New Dictionary(Of String, String)
        mail_dict.Add("mail", "NR") 'NEW REQUEST
        mail_dict.Add("to", syscfg.getSystemAdminMail)
        mail_dict.Add("{rq_id}", req_id.ToString)
        mail_dict.Add("{description}", "<b>" + mailSubject + "</b><br>" + mailBody) 'MAIL SUBJECT / AI DESCRIPTION
        mail_dict.Add("{owners}", ownersMailStr)
		mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")

        newMail.SendNotificationMail(mail_dict)

        '////////////////////////////////////////////////////////////
        'INSERT LOG HERE
        '////////////////////////////////////////////////////////////

        'EVENT: AI_EXTENSION [R5]
        log_dict.Add("request_id", req_id.ToString)
        log_dict.Add("requestor_id", mailFrom)
        log_dict.Add("event", "RQ_CREATED")
        log_dict.Add("detail", "Some detail here...")

        newLog.LogWrite(log_dict)

        dbconn.Close()

    End Sub

    Public Function requestIsUnset(ByVal id As String) As Boolean
        Dim result As Boolean = False
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
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

    Public Sub requestSetDueDate(ByVal id As String)
        Dim result As Boolean = False
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT due FROM requests WHERE id=" & id

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.IsDBNull(0)
        End If

        dbconn.Close()

    End Sub

    Public Sub requestSetStatus(ByVal id As String, ByVal initial As String, ByVal final As String)
        Dim result As Boolean = False
        Dim dbconn As OleDbConnection
        Dim dbcomm_req, dbcomm_req_update As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req, sql_update As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
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

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
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

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
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
        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim link As New Linker
        Dim dbconn As OleDbConnection
        Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
        Dim dbread_req, dbread_ais As OleDbDataReader
        Dim sql, sql_req, sql_ais As String
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


        dbconn = New OleDbConnection(syscfg.getConnection)
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
            mail_dict.Add("to", users.getMailById(requestor))
            mail_dict.Add("{ai_id}", ai_id)
            mail_dict.Add("{owner}", users.getNameById(owner) & "(" & owner & ")")
            mail_dict.Add("{description}", dbread_ais.GetString(2))
            mail_dict.Add("{detail}", dbread_ais.GetString(2))
            mail_dict.Add("{duedate}", dbread_ais.GetDateTime(4).ToString("dd/MMM/yyyy"))
            mail_dict.Add("{delivery}", Now.Date.ToString("dd/MMM/yyyy"))
			mail_dict.Add("{requestor_name}", users.getNameById(requestor))
            mail_dict.Add("{app_link}", syscfg.getSystemUrl)
            mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
            mail_dict.Add("{delivery_link1}", syscfg.getSystemUrl + "delivery.ashx?file=" + fileName1)
            mail_dict.Add("{filename1}", fileName1)

            If fileName2 <> "" Then
                mail_dict.Add("{delivery_link2}", syscfg.getSystemUrl + "delivery.ashx?file=" + fileName2)
                mail_dict.Add("{d2}", "block")
                mail_dict.Add("{filename2}", fileName2)
            Else
                mail_dict.Add("{d2}", "none")
            End If
            If fileName3 <> "" Then
                mail_dict.Add("{delivery_link3}", syscfg.getSystemUrl + "delivery.ashx?file=" + fileName3)
                mail_dict.Add("{d3}", "block")
                mail_dict.Add("{filename3}", fileName3)
            Else
                mail_dict.Add("{d3}", "none")
            End If
            If fileName4 <> "" Then
                mail_dict.Add("{delivery_link4}", syscfg.getSystemUrl + "delivery.ashx?file=" + fileName4)
                mail_dict.Add("{d4}", "block")
                mail_dict.Add("{filename4}", fileName4)
            Else
                mail_dict.Add("{d4}", "none")
            End If
            If fileName5 <> "" Then
                mail_dict.Add("{d5}", "block")
                mail_dict.Add("{delivery_link5}", syscfg.getSystemUrl + "delivery.ashx?file=" + fileName5)
                mail_dict.Add("{filename5}", fileName5)
            Else
                mail_dict.Add("{d5}", "none")
            End If

            mail_dict.Add("{completed_link}", syscfg.getSystemUrl + "sap_completed.aspx?id=" + link.enLink(ai_id.ToString))
            mail_dict.Add("{uncompleted_link}", syscfg.getSystemUrl + "sap_uncompleted.aspx?id=" + link.enLink(ai_id.ToString))

            newMail.SendNotificationMail(mail_dict)

            '////////////////////////////////////////////////////////////
            'INSERT LOG HERE
            '////////////////////////////////////////////////////////////
            'ai_id, request_id, owner_id, requestor_id, admin_id, event, prev_value, new_value, detail

            'EVENT: AI_CREATED [R1]

            Dim newLog As New LogSAPTareas

            Dim log_dict As New Dictionary(Of String, String)
            log_dict.Add("ai_id", ai_id)
            log_dict.Add("request_id", request.ToString)
            log_dict.Add("admin_id", users.getId)
            log_dict.Add("owner_id", owner)
            log_dict.Add("requestor_id", requestor)
            log_dict.Add("new_value", "")
            log_dict.Add("event", "AI_DELIVERED")
            log_dict.Add("detail", fn)

            newLog.LogWrite(log_dict)

        End If

        dbconn.Close()

    End Sub

End Class
