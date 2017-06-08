Imports System.Data.OleDb
Imports System.Text

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
        Dim dbcomm As OleDbCommand
        'Dim dbread_ais As OleDbDataReader
        Dim sql As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New LogSAPTareas

        'ID LAST REQUEST CREATED
        Dim req_id As Long
        Dim ai_id As Long

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        Dim mailFrom As String = sourceMail("from")
        Dim mailSubject As String = sourceMail("subject")
        Dim mailBody As String = sourceMail("body")
        Dim mailDate As String = sourceMail("date")
        Dim mailID As String = sourceMail("id")
        Dim mailRaw As String = sourceMail("raw")
        Dim requestorID As String = users.getIdByMail(mailFrom)

        Dim auxMailDate As Date = Date.Parse(mailDate)
        mailDate = auxMailDate.ToString("yyyy-MM-dd hh:mm:ss")

        Dim dueString As String = mailBody.ToLower
        Dim posDueDate As Date
        Dim parseResult As Boolean = False
        Dim dateFormats() As String = {"dd/MMM/yyyy", "d/MMM/yyyy", "MMM/dd/yyyy", "MMM/d/yyyy", "dd-MMM-yyyy", "d-MMM-yyyy", "MMM-dd-yyyy", "MMM-d-yyyy", "dd MMM yyyy", "d MMM yyyy", "MMM dd yyyy", "MMM d yyyy", "dd_MMM_yyyy", "d_MMM_yyyy", "MMM_dd_yyyy", "MMM_d_yyyy", "ddMMMyyyy", "dMMMyyyy", "MMMddyyyy", "MMMdyyyy"}

        Dim dueIndex As Integer = dueString.IndexOf("due date")
        If dueIndex = -1 Then
            dueIndex = dueString.IndexOf("dd")
            If dueIndex = -1 Then
                dueIndex = dueString.IndexOf("due")
            End If
        End If

        Dim dueStringArray() As String = dueString.Split(Environment.NewLine)

        'FIND DUE DATE
        If dueIndex <> -1 Then
            For Each dueLine In dueStringArray

                If dueLine.IndexOf("due date") <> -1 Or dueLine.IndexOf("due") <> -1 Or dueLine.IndexOf("dd") <> -1 Then

                    dueLine = dueLine.Replace("due", "").Trim
                    dueLine = dueLine.Replace("dd", "").Trim
                    dueLine = dueLine.Replace("date", "").Trim
                    dueLine = dueLine.Replace(":", "").Trim

                    parseResult = Date.TryParseExact(dueLine, dateFormats, New System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.None, posDueDate)

                    'if parsing due date is true, then exit loop
                    If parseResult Then
                        Exit For
                    End If

                End If
            Next
        End If

        'CREATE REQUEST
        'IF THE MAIL HAS AT LEAST ONE OWNER THEN SET STATUS TO IN PROGRESS
        'ELSE THE RQ STATUS IS PENDING...
        Dim reqStatus As String = "PD"
        If sourceMail.Count > 6 Then
            'reqStatus = "IP"
        End If

        'TRUNCATE DESCRIPTION
        If mailSubject.Length > 250 Then
            mailSubject = (Mid(mailSubject, 1, 250) & "...").Trim()
        End If

        If mailBody.Length > 250 Then
            mailBody = (Mid(mailBody, 1, 250) & "...").Trim()
        End If

        Dim dueSQLfield As String = ""
        Dim dueSQLdata As String = ""
        If parseResult Then
            dueSQLfield = ", due"
            dueSQLdata = ", '" + posDueDate.Year.ToString + "-" + posDueDate.Month.ToString + "-" + posDueDate.Day.ToString + "'"
            reqStatus = "CR"
        End If

        sql = "INSERT INTO requests (requestor, requested, mail, subject, detail, status, mail_id" + dueSQLfield + ") VALUES ('" + requestorID + "', '" + mailDate + "', '" + mailRaw + "', '" + mailSubject + "', '" + mailBody + "', '" + reqStatus + "', '" + mailID + "'" + dueSQLdata + ")"

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
        Dim adminMail As String = ""

        For i = 1 To sourceMail.Count - 6 'THE MAIL HAS 6 FIELDS THE REST ARE THE OWNERS

            owner = "own" + i.ToString
            ownerID = users.getIdByMail(sourceMail(owner))

            'CREATE ACTION ITEM
            sql = "INSERT INTO actionitems (request_id, description, owner, status" + dueSQLfield + ") VALUES (" + req_id.ToString + ", '" + mailSubject + "', '" + ownerID + "', 'PD'" + dueSQLdata + ");" '+ " / " + mailBody

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
            mail_dict.Add("{ai_owner}", users.getNameById(ownerID))
            mail_dict.Add("{app_link}", syscfg.getSystemUrl)
            mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
            mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + ai_id.ToString)

            'DO NOT SEND BECAUSE THE AI DOES NOT HAVE DUE DATE TO ACCEPT
            If parseResult Then
                newMail.SendNotificationMail(mail_dict)
            End If

            log_dict.Add("request_id", req_id.ToString)
            log_dict.Add("ai_id", ai_id.ToString)
            log_dict.Add("requestor_id", requestorID)
            log_dict.Add("admin_id", "system")
            log_dict.Add("owner_id", ownerID)
            log_dict.Add("event", "AI_CREATED")
            log_dict.Add("detail", "")

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
        mail_dict.Add("{rq_link}", syscfg.getSystemUrl & "sap_req.aspx?id=" & req_id.ToString.Trim())
        mail_dict.Add("{description}", "<b>" + mailSubject + "</b><br>" + mailBody) 'MAIL SUBJECT / AI DESCRIPTION
        mail_dict.Add("{owners}", ownersMailStr)
        mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        mail_dict.Add("{contact_mail_link}", "mailto: " & users.getAdminMail & "?subject=Questions about the report")

        'SEND TO MULTIPLE REQUESTORS
        For Each adminMail In syscfg.getSystemAdminMail.Split(";")
            If Not String.IsNullOrEmpty(adminMail) Then
                mail_dict("to") = adminMail
                newMail.SendNotificationMail(mail_dict)
            End If
        Next

        '////////////////////////////////////////////////////////////
        'INSERT LOG HERE
        '////////////////////////////////////////////////////////////

        'EVENT: AI_EXTENSION [R5]
        log_dict.Add("request_id", req_id.ToString)
        log_dict.Add("requestor_id", requestorID)
        log_dict.Add("event", "RQ_CREATED")
        log_dict.Add("detail", "")

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

    Public Function requestMailProcessed(ByVal id As String) As Boolean
        Dim result As Boolean = False
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT * FROM requests WHERE mail_id='" & id & "'"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            result = True
        End If

        dbconn.Close()

        Return result

    End Function


    Private Function ais_sum_filter(ByVal filter As String) As Integer
        Dim syscfg As New SysConfig
        Dim actions As New SapActions
        Dim dbconn As OleDbConnection
        Dim dbcomm As OleDbCommand
        Dim dbread As OleDbDataReader

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        Dim extra_where As String = "status <> 'CP' AND status <> 'XX'"
        Dim extra_subq As String = ""

        Select Case filter

            Case "ur"
                extra_where = "ai_count <= 0"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id) AS ai_count"

            Case "nd"
                extra_where = "(status = 'ND')"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'ND') AS ai_count"

            Case "ap"
                extra_where = "(status = 'PD')"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'PD' AND actionitems.owner IS NOT null) AS ai_count"

            Case "rq"
                extra_where = "(status = 'PD' AND ai_count > 0)"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'DL' AND sent_confirm = 1) AS ai_count"

            Case "du"
                extra_where = "(status <> 'DL' AND status <> 'PD') AND ai_count > 0"
                extra_subq = ", (SELECT COUNT(distinct actionitems.owner) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.owner IS NOT null GROUP BY actionitems.request_id HAVING COUNT(distinct actionitems.owner) > 1 ) AS ai_count"

            Case "ex"
                extra_where = "(status = 'NE') OR (ai_count > 0) "
                extra_subq = ", (SELECT Count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'NE') AS ai_count"

            Case "pr"
                extra_where = "(status = 'DL')"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'DL') AS ai_count"

            Case "dw"
                extra_where = "((status <> 'DL' OR status <> 'PD') AND due IS NOT null AND DATEDIFF(day, TODAY(), due) >= 0)"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND (actionitems.status <> 'DL' OR actionitems.status <> 'PD') AND (DATEDIFF(day, actionitems.due, TODAY())) >= 0 AND actionitems.owner IS NOT null) AS ai_count"

            Case "ov"
                extra_where = "(status <> 'DL' OR status <> 'PD') AND due IS NOT null AND ai_count >= 0"
                extra_subq = ", (SELECT COUNT(distinct actionitems.id) FROM actionitems WHERE requests.id = actionitems.request_id AND (actionitems.status <> 'DL' AND actionitems.status <> 'PD') AND (DATEDIFF(day, TODAY(), actionitems.due)) < 0 AND actionitems.owner IS NOT null) AS ai_count"

        End Select

        Dim sql As String = "SELECT * " & extra_subq & " FROM requests WHERE " & extra_where & " ORDER BY id DESC"
        Dim res As Integer = 0
        dbcomm = New OleDbCommand(sql, dbconn)
        dbread = dbcomm.ExecuteReader()

        If dbread.HasRows Then

            While dbread.Read()
                Dim current_req_id As Long = Convert.ToInt64(dbread.GetValue(0))
                Dim req_duedate As Date
                Dim req_missing_days As Integer = 0

                If Not dbread.IsDBNull(6) Then
                    req_duedate = dbread.GetDateTime(6)
                    req_missing_days = DateDiff(DateInterval.Day, Today.Date, req_duedate.Date)
                Else
                    req_missing_days = -1
                    req_duedate = Today()
                End If

                If filter <> "od" Or (filter = "od" And req_missing_days < 7 And req_missing_days > -1) Then
                    res = res + 1
                End If

            End While

        End If

        dbconn.Close()
        Return res

    End Function

    Public Function getDataAdminReport() As String

        Dim todayList As String = ""
        Dim count As Integer
        Dim syscfg As New SysConfig

        'Building Owner
        'Count Acceptance Pending OW
        count = ais_sum_filter("ap")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Owner</td> 
                                    <td style='width: 200px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'><a href='" & syscfg.getSystemUrl & "sap_main.aspx?f=ap'>Acceptance Pending (OW)</a></td>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif;'>" & count & "</td>
                                    <td style='width: 330px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Get feedback and accept/reject the request in the system</td>
                                </tr>"
        'Count overdue items
        count = ais_sum_filter("ov")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'></td> 
                                    <td style='width: 200px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'><a href='" & syscfg.getSystemUrl & "sap_main.aspx?f=ov'>Overdue</a></td>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif;'>" & count & "</td>
                                    <td style='width: 330px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Get feedback about the Status, brief the Requestor and update the information in the system</td>
                                </tr>"
        ''Building Requestor
        'Count Acceptance Pending RQ
        count = ais_sum_filter("rq")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Requestor</td> 
                                    <td style='width: 200px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'><a href='" & syscfg.getSystemUrl & "sap_main.aspx?f=rq'>Acceptance Pending (RQ)</a></td>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif;'>" & count & "</td>
                                    <td style='width: 330px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Ask the owner to accept it</td>
                                </tr>"

        'Count extension pending items
        count = ais_sum_filter("ex")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'></td> 
                                    <td style='width: 200px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'><a href='" & syscfg.getSystemUrl & "sap_main.aspx?f=ex'>Extension Pending</a></td>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif;'>" & count & "</td>
                                    <td style='width: 330px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Get feedback and accept/reject the request in the system</td>
                                </tr>"
        'Count unassigned items
        count = ais_sum_filter("ur")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'></td> 
                                    <td style='width: 200px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'><a href='" & syscfg.getSystemUrl & "sap_main.aspx?f=ur'>Unassigned</a></td>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif;'>" & count & "</td>
                                    <td style='width: 330px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Ask about the owner of the request and update the info in the system</td>
                                </tr>"
        'Count need data items
        count = ais_sum_filter("nd")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'></td> 
                                    <td style='width: 200px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'><a href='" & syscfg.getSystemUrl & "sap_main.aspx?f=nd'>Need Data</a></td>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif;'>" & count & "</td>
                                    <td style='width: 330px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Ask about the missing info and update it in the system</td>
                                </tr>"
        ''Building Admin
        'Count this week items
        count = ais_sum_filter("dw")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Admin</td> 
                                    <td style='width: 200px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'><a href='" & syscfg.getSystemUrl & "sap_main.aspx?f=dw'> This Week DD</a></td>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif;'>" & count & "</td>
                                    <td style='width: 330px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Be on top of this week deliverables</td>
                                </tr>"
        'Count multiple owners items
        count = ais_sum_filter("du")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'></td> 
                                    <td style='width: 200px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'><a href='" & syscfg.getSystemUrl & "sap_main.aspx?f=du'> Multiple Owner</a></td>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif;'>" & count & "</td>
                                    <td style='width: 330px;border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>This are more complex projects, you might want to monitor closelly to make sure that everything goes as expected</td>
                                </tr>"

        Return todayList

    End Function

    Public Function getDataOwnerReport(ByVal id As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String
        Dim sqlOwner As New StringBuilder

        Dim syscfg As New SysConfig
        Dim users As New SapUser

        Dim todayList As String = ""
        Dim weekList As String = ""

        Dim dueDate As Date
        Dim aiStatus As String
        Dim daysDiff As Integer
        Dim daysDiffStr As String = ""
        Dim dueStr As String = ""
        Dim aiID As Integer
        Dim aiDesc As String
        Dim newLine As String

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'sql_req = "SELECT * FROM actionitems WHERE owner='" & id & "' AND status <> 'CP' AND status <> 'XX'"

        sqlOwner.Append("SELECT * FROM actionitems WHERE owner='" & id & "' AND status = 'PD' ")
        sqlOwner.Append(" UNION ")
        sqlOwner.Append("SELECT * FROM actionitems WHERE owner='" & id & "' AND status IN ('DL', 'NE', 'OD') ")
        sqlOwner.Append(" UNION ")
        sqlOwner.Append("SELECT * FROM actionitems WHERE owner='" & id & "' AND status = 'IP' ")
        sqlOwner.Append("ORDER BY due ASC")
        sql_req = sqlOwner.ToString

        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            While dbread_req.Read()

                aiStatus = dbread_req.GetString(5)
                If dbread_req.IsDBNull(4) Then
                    daysDiffStr = "Unset"
                    dueStr = "Unset"
                Else
                    dueDate = dbread_req.GetDateTime(4)
                    daysDiff = Math.Abs(DateDiff("d", Today.Date, dueDate))
                    daysDiffStr = daysDiff.ToString
                    dueStr = dueDate.ToString("dd/MMM/yyyy")
                    If dueDate < Today.Date Then
                        aiStatus = "OD"
                    End If
                End If
                aiID = dbread_req.GetInt64(0)
                aiDesc = dbread_req.GetString(2)

                'newLine = Environment.NewLine & "<tr><td> <a href='" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & "'>#" & aiID.ToString & "</a></td><td>"

                Select Case aiStatus

                    Case "PD"
                        newLine = Environment.NewLine & "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Pending</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>New Action Item</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"

                        todayList = todayList + newLine
                        weekList = weekList + newLine

                    Case "DL"
                        newLine = Environment.NewLine & "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Delivered</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Already delivered</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
                        todayList = todayList + newLine
                        weekList = weekList + newLine


                    Case "NE"
                        newLine = Environment.NewLine & "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Need Extension</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Extension required</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
                        todayList = todayList + newLine
                        weekList = weekList + newLine

                    Case "OD"
                        newLine = Environment.NewLine & "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Overdue</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Overdue " & daysDiffStr & " day</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
                        todayList = todayList + newLine
                        weekList = weekList + newLine

                    Case "IP"
                        newLine = Environment.NewLine & "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>In Progress</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>AI in Progress</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
                        todayList = todayList + newLine
                        weekList = weekList + newLine

                    Case Else
                        Select Case daysDiffStr
                            Case "0"
                                newLine = Environment.NewLine &
                                                        "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>-</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Delivery Today</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
                                todayList = todayList + newLine
                            Case "2"
                                newLine = Environment.NewLine &
                                                    "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>-</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Confirm Today</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
                                todayList = todayList + newLine
                            Case Else
                                If daysDiff < 6 Then
                                    newLine = Environment.NewLine &
                                                    "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>-</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Delivery this week</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
                                    weekList = weekList + newLine
                                Else
                                    newLine = Environment.NewLine &
                                                    "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>-</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Delivery after this week</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
                                    weekList = weekList + newLine
                                End If
                        End Select
                End Select

            End While
        End If

        'If String.IsNullOrEmpty(todayList) Then
        '    todayList = "Nothing pending for Today"
        'End If

        'If String.IsNullOrEmpty(weekList) Then
        '    weekList = "Nothing pending for This Week"
        'End If

        dbconn.Close()

        result.Add("today", todayList)
        result.Add("this week", weekList)
        Return result

    End Function

    Public Sub sendOwnerReport(ByVal interval As String)
        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim newMail As New MailTemplate
        Dim id As String
        Dim report As Dictionary(Of String, String)
        Dim mail_dict As New Dictionary(Of String, String)

        For Each id In users.getOwnersID("")

            report = getDataOwnerReport(id)

            'IF THERE ARE ANY NEWS
            If Not String.IsNullOrEmpty(report(interval)) Then

                mail_dict.Add("mail", "OR") 'OWNER REPORT
                mail_dict.Add("to", users.getMailById(id))
                mail_dict.Add("{time}", interval)
                mail_dict.Add("{date}", DateTime.Now.ToString("MM/dd/yyyy"))
                mail_dict.Add("{data}", report(interval))
                mail_dict.Add("{app_link}", syscfg.getSystemUrl)
                mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the Owner report")

                newMail.SendNotificationMail(mail_dict)

                mail_dict.Clear()
            End If
        Next

    End Sub

    Public Sub sendAdminReport()
        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim newMail As New MailTemplate
        Dim id As String
        Dim mail_dict As New Dictionary(Of String, String)
        Dim report As String

        report = getDataAdminReport()

        If report = "" Then
            report = "No action is required"
        End If

        For Each id In users.getAdminsID("")

            'IF THERE ARE ANY NEWS
            If Not String.IsNullOrEmpty(report) Then

                mail_dict.Add("mail", "AR") 'ADMIN REPORT
                mail_dict.Add("to", users.getMailById(id))
                mail_dict.Add("{time}", "Today")
                mail_dict.Add("{date}", DateTime.Now.ToString("MM/dd/yyyy"))
                mail_dict.Add("{data}", report)
                mail_dict.Add("{app_link}", syscfg.getSystemUrl)
                mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the Admin report")

                newMail.SendNotificationMail(mail_dict)

                mail_dict.Clear()
            End If
        Next
    End Sub


End Class
