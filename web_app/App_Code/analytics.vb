﻿Imports Microsoft.VisualBasic

Imports System.Data.OleDb
Imports System
Imports SapActions

Public Class SapAnalytics

    Public each_section_width As Integer = 100

    Public Function createTimeLine(ByVal req_id As Integer) As String
        'CHECK IF DB EXISTS
        '#####TODO###################
        Dim dbconn As OleDbConnection
        Dim dbcomm_req, dbcomm_req_log As OleDbCommand
        Dim dbread_req, dbread_req_log As OleDbDataReader
        Dim sql_req, sql_req_log As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New LogSAPTareas
        Dim actions As New SapActions

        Dim evento As String
        Dim result As String = ""
        Dim line, color As String
        Dim timestamp As Date
        Dim creation, extension As Date
        Dim duedate As Date = actions.requestGetAIsLastDueDate(req_id)
        Dim startyear As Integer = 0
        Dim desde As String = ""
        Dim timestr As String = ""
        Dim hasta As Date
        Dim caso As String
        Dim dayPxls, duration, margin As Integer

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'GET THE REQUEST
        sql_req = "SELECT * FROM requests WHERE id=" & req_id.ToString & ";"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()

        'IF THE REQUEST EXISTS
        If dbread_req.HasRows Then
            dbread_req.Read()
            'REPRESENT FROM CREATION DATE TO DUE DATE PLUS EXTENSION DATE

            'BASED ON THE COUNT OF DAYS
            creation = dbread_req.GetDateTime(2)
            'duedate = dbread_req.GetDateTime(0)
            'extension = dbread_req.GetDateTime(0)

            hasta = Today.Date
            dayPxls = Int(600 / (Math.Abs(DateDiff("d", creation, hasta)) + 1))

            If duedate <> Nothing Then
                If duedate < Today Then
                    caso = "OVERDUE"
                    duration = Math.Abs(DateDiff("d", creation, duedate))
                    margin = Int(dayPxls * duration)
                    line = "<span style=""margin-left: 0px; width: " + Int(dayPxls * duration).ToString + "px;"" class=""bubble bubble-dolor"" data-duration=""" + duration.ToString + """></span>"
                    duration = Math.Abs(DateDiff("d", duedate, Today.Date)) + 1
                    line = line + "<span style=""margin-left: 0px; width: " + Int(dayPxls * duration).ToString + "px;"" class=""bubble bubble-default"" data-duration=""" + duration.ToString + """></span>"
                    line = line + "<span class=""date"" style=""display: inline;"">" + duedate.ToString("dddd, MMMM dd, yyyy") + "</span> <span class=""label"" style=""display: inline;""> <b>Overdue " + duration.ToString + " days</b></span>"
                Else
                    caso = "ON SCHEDULE"
                    hasta = duedate
                    dayPxls = Int(600 / (Math.Abs(DateDiff("d", creation, duedate)) + 1))
                    duration = Math.Abs(DateDiff("d", creation, Today.Date))
                    line = "<span style=""display: inline; margin-left: 0px; width: " + Int(dayPxls * duration).ToString + "px;"" class=""bubble bubble-lorem"" data-duration=""" + duration.ToString + """></span>"
                    duration = Math.Abs(DateDiff("d", duedate, Today.Date))
                    line = line + "<span style=""margin-left: 0px; width: " + Int(dayPxls * duration).ToString + "px;"" class=""bubble bubble-ipsum"" data-duration=""" + duration.ToString + """></span>"
                    line = line + "<span class=""date"" style=""display: inline;"">" + duedate.ToString("dddd, MMMM dd, yyyy") + "</span> <span class=""label"" style=""display: inline;""> <b>On Schedule " + duration.ToString + " days left</b></span>"
                End If
            Else
                caso = "NO DUEDATE"
                duration = Math.Abs(DateDiff("d", creation, Today.Date))
                line = line + "<span style=""margin-left: 0px; width: " + Int(dayPxls * duration).ToString + "px;"" class=""bubble bubble-dolor"" data-duration=""" + duration.ToString + """></span>"
                line = line + "<span class=""date"">" + creation.ToString("dddd, MMMM dd, yyyy") + "</span> <span class=""label"" style=""display: inline;""> <b>No Due Date " + duration.ToString + " days since</b></span>"
            End If
            line = "<li>" + line + "</li>"

            result = result + line

            'SHOW TODAY
            line = ""
            timestamp = Today.Date
            duration = Math.Abs(DateDiff("d", creation.Date, timestamp) - 1)
            line = line + "<span style=""margin-left: " + Int(dayPxls * duration).ToString + "px; width: " + Int(dayPxls).ToString + "px;"" class=""bubble bubble-" + "dolor" + """ data-duration=""" + duration.ToString + """></span>"
            duration = 1
            line = line + "<span class=""date"">" + timestamp.ToString("dddd, MMMM dd, yyyy") + "</span> <span class=""label"" style=""display: inline;""> <b>" + "Today" + "</b></span>"
            line = "<li>" + line + "</li>"
            result = line + result

        End If

        'GET THE REQUEST LOG
        sql_req_log = "SELECT * FROM log WHERE request_id=" & req_id.ToString & " ORDER BY time_stamp DESC;"
        dbcomm_req_log = New OleDbCommand(sql_req_log, dbconn)
        dbread_req_log = dbcomm_req_log.ExecuteReader()

        If dbread_req_log.HasRows Then
            While dbread_req_log.Read()

                Select Case dbread_req_log.GetString(7)
                    Case "AI_CREATED"
                        color = "ipsum"
                        evento = "AI Created"
                    Case "AI_ACCEPTED"
                        color = "lorem"
                        evento = "AI Accepted"
                    Case "AI_CONFIRMED"
                        color = "dolor"
                        evento = "AI Confirmed"
                    Case "AI_EXTENSION"
                        color = "default"
                        evento = "AI Extension Required"
                    Case "AI_EXTENDED"
                        color = "ipsum"
                        evento = "AI Extended"
                    Case "AI_NOT_EXTENDED"
                        color = "default"
                        evento = "AI Extension Rejected"
                    Case "AI_UPDATED"
                        color = "default"
                        evento = "AI Updated"
                    Case "AI_DELIVERED"
                        color = "default"
                        evento = "AI Updated"
                    Case "AI_CANCELED"
                        color = "default"
                        evento = "AI Canceled"
                    Case "RQ_CREATED"
                        color = "default"
                        evento = "RQ Created"
                    Case "RQ_CANCELED"
                        color = "default"
                        evento = "RQ Canceled"
                    Case "RQ_UPDATED"
                        color = "default"
                        evento = "RQ Updated"
                    Case "MISSING_INFO"
                        color = "ipsum"
                        evento = "RQ Missing Info"
                    Case Else
                        color = "ipsum"
                        evento = "N/A"
                End Select

                line = ""
                timestamp = dbread_req_log.GetDateTime(1).Date
                duration = Math.Abs(DateDiff("d", creation.Date, timestamp))
                If duration > 1 Then
                    duration = duration - 1
                End If
                line = line + "<span style=""display: inline; margin-left: " + Int(dayPxls * duration).ToString + "px; width: " + Int(dayPxls).ToString + "px;"" class=""bubble bubble-" + color + """ data-duration=""" + duration.ToString + """></span>"
                duration = 1
                line = line + "<span class=""date"" style=""display: inline;"">" + timestamp.ToString("dddd, MMMM dd, yyyy") + "</span> <span class=""label"" style=""display: inline;""> <b>" + evento + "</b></span>"
                line = "<li>" + line + "</li>"
                result = result + line
            End While

        End If

        dbconn.Close()

        Dim sections As String = ""
        Dim section_name As String = ""
        Dim section_tail As String = ""
        Dim section_width As Integer = 0
        Dim ndx As Integer

        duration = Math.Abs(DateDiff("d", creation, hasta))

        'SHOW A WEEK OR LESS
        Select Case duration
            Case Is < 8
                section_name = "Day"
                section_width = 1
                section_tail = "Today"
            Case Is < 40
                section_name = "Week"
                section_width = 7
                section_tail = "This Week"
            Case Is < 200
                section_name = "Month"
                section_width = 30
                section_tail = "This Month"
            Case Else
                section_name = "Year"
                section_width = 365
                section_tail = "This Year"
        End Select

        For ndx = 1 To Int(duration / section_width)
            sections = sections + "<section>" + section_name + " " + ndx.ToString + "</section>"
        Next

        sections = sections + "<section>" + section_tail + "</section>"

        result = "<div class=""scale"">" + sections + "</div><ul class=""data"">" + result + "</ul></div>"

        each_section_width = section_width * dayPxls

        Return result

    End Function


    Public Function createRequestsReport(ByVal fromDate As Date, ByVal toDate As Date) As String
        'CHECK IF DB EXISTS
        '#####TODO###################
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New LogSAPTareas
        Dim actions As New SapActions

        Dim tr As String
        Dim result As String = ""
        Dim startyear As Integer = 0
        Dim desde As String = ""
        Dim timestr As String = ""

        Dim sb As New StringBuilder()

        Dim requestID As Integer
        Dim requestedDateStr As String
        Dim dueDateStr As String
        Dim completedDateStr As String
        Dim needDataCountStr As String
        Dim createdDate As Date
        Dim createdDatestr As String
        Dim inProgressDate As Date
        Dim inProgressDateStr As String

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'GET THE REQUEST
        sql_req = "SELECT * FROM requests WHERE status='CP' AND due >= '" & fromDate.ToString("yyyy-MM-dd") & "' AND due <= '" & toDate.ToString("yyyy-MM-dd") & "';"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()

        sb.AppendLine("<table id='report' class='tablesorter'><thead><tr><th class='header'>Requestor</th><th class='header'>RQ#</th><th class='header'>Subject</th><th class='header headerSortDown'>Requested</th><th class='header'>Need Data</th><th class='header'>Created</th><th class='header'>In Progress</th><th class='header'>Due</th><th class='header'>Completed</th></tr></thead><tbody>")

        'IF THE REQUEST EXISTS
        If dbread_req.HasRows Then

            While dbread_req.Read()

                requestedDateStr = "Unset"
                dueDateStr = "Unset"
                completedDateStr = "Unset"

                If Not dbread_req.IsDBNull(2) Then
                    requestedDateStr = dbread_req.GetDateTime(2).ToString("dd/MMM/yyyy")
                End If

                If Not dbread_req.IsDBNull(6) Then
                    dueDateStr = dbread_req.GetDateTime(6).ToString("dd/MMM/yyyy")
                End If

                If Not dbread_req.IsDBNull(10) Then
                    completedDateStr = dbread_req.GetDateTime(10).ToString("dd/MMM/yyyy")
                End If

                requestID = dbread_req.GetInt64(0)
                needDataCountStr = actions.getCountFromLog("request_id", requestID, "MISSING_INFO")
                If actions.getFirstDateFromLog("request_id", requestID, "AI_CREATED", createdDate) Then
                    createdDatestr = createdDate.ToString("dd/MMM/yyyy")
                Else
                    createdDatestr = "Unset"
                End If
                If actions.getFirstDateFromLog("request_id", requestID, "AI_ACCEPTED", inProgressDate) Then
                    inProgressDateStr = inProgressDate.ToString("dd/MMM/yyyy")
                Else
                    inProgressDateStr = "Unset"
                End If

                tr = ""
                tr = tr & "<td>" & dbread_req.GetString(1) & "</td>"
                tr = tr & "<td>" & requestID.ToString & "</td>"
                tr = tr & "<td>" & dbread_req.GetString(4) & "</td>"
                tr = tr & "<td>" & requestedDateStr & "</td>"
                tr = tr & "<td>" & needDataCountStr & "</td>"
                tr = tr & "<td>" & createdDatestr & "</td>"
                tr = tr & "<td>" & inProgressDateStr & "</td>"
                tr = tr & "<td>" & dueDateStr & "</td>"
                tr = tr & "<td>" & completedDateStr & "</td>"
                tr = "<tr>" & tr & "</tr>"
                sb.AppendLine(tr)

            End While

        End If

        sb.AppendLine("</tbody></body></html>")

        dbconn.Close()

        Return sb.ToString

    End Function


    Public Function createActionItemsReport(ByVal fromDate As Date, ByVal toDate As Date) As String
        'CHECK IF DB EXISTS
        '#####TODO###################
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig
        Dim users As New SapUser
        Dim eLink As New Linker
        Dim newLog As New LogSAPTareas
        Dim actions As New SapActions

        Dim tr As String
        Dim result As String = ""
        Dim startyear As Integer = 0
        Dim desde As String = ""
        Dim timestr As String = ""

        Dim sb As New StringBuilder()

        Dim requestID As Integer
        Dim actionitemID As Integer
        Dim deliveredDateStr As String
        Dim dueDateStr As String
        Dim completedDateStr As String
        Dim originalDateStr As String
        Dim extensionCountStr As String
        Dim deliveredCountStr As String
        Dim acceptedDate As Date
        Dim acceptedDatestr As String

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        'GET THE REQUEST
        sql_req = "SELECT * FROM actionitems WHERE status='CP' AND due >= '" & fromDate.ToString("yyyy-MM-dd") & "' AND due <= '" & toDate.ToString("yyyy-MM-dd") & "' ORDER BY owner;"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()

        sb.AppendLine("<table id='report' class='tablesorter'><thead><tr><th class='header'>Owner</th><th class='header'>AI#</th><th class='header'>RQ#</th><th class='header'>Description</th><th class='header'>Created</th><th class='header'>Extension</th><th class='header'>In Progress</th><th class='header'>Delivered</th><th class='header'>Delivered</th><th class='header'>Due</th><th class='header'>Original</th><th class='header'>Completed</th></tr></thead><tbody>")

        'IF THE REQUEST EXISTS
        If dbread_req.HasRows Then

            While dbread_req.Read()

                dueDateStr = "Unset"
                deliveredDatestr = "Unset"
                originalDateStr = "Unset"
                completedDateStr = "Unset"

                If Not dbread_req.IsDBNull(4) Then
                    dueDateStr = dbread_req.GetDateTime(4).ToString("dd/MMM/yyyy")
                End If

                If Not dbread_req.IsDBNull(10) Then
                    originalDateStr = dbread_req.GetDateTime(10).ToString("dd/MMM/yyyy")
                End If

                If Not dbread_req.IsDBNull(13) Then
                    deliveredDateStr = dbread_req.GetDateTime(13).ToString("dd/MMM/yyyy")
                End If

                If Not dbread_req.IsDBNull(16) Then
                    completedDateStr = dbread_req.GetDateTime(16).ToString("dd/MMM/yyyy")
                End If

                actionitemID = dbread_req.GetInt64(0)
                requestID = dbread_req.GetInt64(1)

                extensionCountStr = actions.getCountFromLog("ai_id", actionitemID, "AI_EXTENSION")
                deliveredCountStr = actions.getCountFromLog("ai_id", actionitemID, "AI_DELIVERED")
                If actions.getFirstDateFromLog("ai_id", actionitemID, "AI_ACCEPTED", acceptedDate) Then
                    acceptedDatestr = acceptedDate.ToString("dd/MMM/yyyy")
                Else
                    acceptedDatestr = "Unset"
                End If

                tr = ""
                tr = tr & "<td>" & dbread_req.GetString(6) & "</td>"
                tr = tr & "<td>" & actionitemID.ToString & "</td>"
                tr = tr & "<td>" & requestID.ToString & "</td>"
                tr = tr & "<td>" & dbread_req.GetString(2) & "</td>"
                tr = tr & "<td>" & dbread_req.GetDateTime(3).ToString("dd/MMM/yyyy") & "</td>"
                tr = tr & "<td>" & extensionCountStr & "</td>"
                tr = tr & "<td>" & acceptedDatestr & "</td>"
                tr = tr & "<td>" & deliveredCountStr & "</td>"
                tr = tr & "<td>" & deliveredDateStr & "</td>"
                tr = tr & "<td>" & dueDateStr & "</td>"
                tr = tr & "<td>" & originalDateStr & "</td>"
                tr = tr & "<td>" & completedDateStr & "</td>"
                tr = "<tr>" & tr & "</tr>"
                sb.AppendLine(tr)

            End While

        End If

        sb.AppendLine("</tbody></body></html>")

        dbconn.Close()

        Return sb.ToString

    End Function


End Class
