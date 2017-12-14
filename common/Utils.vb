Imports System.Net
Imports System.Web

Public Class Utils

    Public Function humanize_Fwd(ByVal i As Integer) As String
        Dim result As String
        Select Case i
            Case Is < 0
                result = "Overdue"
            Case 0
                result = "Today"
            Case 1
                result = "Tomorrow"
            Case Else
                result = "within " + i.ToString + " days"
        End Select
        Return result

    End Function

    Public Function humanize_Bkw(ByVal i As Integer) As String
        Dim result As String
        Select Case i
            Case Is < 0
                result = "Error"
            Case 0
                result = "Today"
            Case 1
                result = "Yesterday"
            Case Else
                result = i.ToString + " days ago"
        End Select
        Return result
    End Function

    Public Function ai_Str_Status(ByVal s As String) As String
        Dim result As String
        Select Case s
            Case "CR"
                result = "Created"
            Case "CP"
                result = "Completed"
            Case "PD"
                result = "Pending"
            Case "IP"
                result = "In Progress"
            Case "ND"
                result = "Need Data"
            Case "NE"
                result = "Extending"
            Case "OD"
                result = "Overdue"
            Case "CF"
                result = "Confirmed"
            Case "DL"
                result = "Delivered"
            Case Else
                result = "Unset"
        End Select
        Return result
    End Function

    Public Function formatDateToSTring(ByVal dateToFormat As Date) As String
        Dim formatedDate As String = String.Empty

        formatedDate = dateToFormat.ToString("dd/MMM/yyyy")

        Return formatedDate
    End Function

    Public Function getEmailTemplate(ByVal mailAcronym) As String
        Dim emailTemplate As String = String.Empty

        Select Case mailAcronym
            Case "ND"
                emailTemplate = ".\email-templates\SAP Email A - More info.html"
            Case "CF"
                emailTemplate = ".\email-templates\SAP Email B - Due 2 days.html"
            Case "CR"
                emailTemplate = ".\email-templates\SAP Email C - AI Owner.html"
            Case "CP"
                emailTemplate = ".\email-templates/SAP Email D - Delivery Approval.html"
            Case "NE"
                emailTemplate = ".\email-templates\SAP Email E - Extension Requested.html"
            Case "EA"
                emailTemplate = ".\email-templates\SAP Email F - Extension Approved.html"
            Case "ER"
                emailTemplate = ".\email-templates\SAP Email G - Extension Rejected.html"
            Case "DL"
                emailTemplate = ".\email-templates\SAP Email I - Due today.html"
            Case "NR"
                emailTemplate = ".\email-templates\SAP Email H - Admin New Request.html"
            Case "OR"
                emailTemplate = ".\email-templates\SAP Email N - Owner Report.html"
            Case "AR"
                emailTemplate = ".\email-templates\SAP Email O - Admin Report.html"
            Case "AC"
                emailTemplate = ".\email-templates/SAP Email J - AI Completed.html"
            Case "AU"
                emailTemplate = ".\email-templates/SAP Email K - AI Uncompleted.html"
            Case "OC"
                emailTemplate = ".\email-templates/SAP Email P - AI Owner Changed.html"
            Case "ET"
                emailTemplate = ".\email-templates/emailTemplate.html"
            Case Else
                emailTemplate = String.Empty
        End Select

        Return emailTemplate
    End Function

    Public Function getDefaultEmailSubject(ByVal mailAcronym) As String
        Dim mailSubject As String = String.Empty

        Select Case mailAcronym
            Case "ND"
                mailSubject = "Your request is pending for information lack"
            Case "CF"
                mailSubject = "AI due date confirmation"
            Case "CR"
                mailSubject = "New AI"
            Case "CP"
                mailSubject = "Delivery approval"
            Case "NE"
                mailSubject = "AI extension requested"
            Case "EA"
                mailSubject = "AI extension approved"
            Case "ER"
                mailSubject = "AI extension rejected"
            Case "DL"
                mailSubject = "AI delivery day"
            Case "NR"
                mailSubject = "New RQ created"
            Case "OR"
                mailSubject = "Your WIP Action Items"
            Case "AR"
                mailSubject = "COO Project Tracking - Admin Report"
            Case "AC"
                mailSubject = "AI completed"
            Case "AU"
                mailSubject = "AI uncompleted"
            Case Else
                'ERROR / HALT DO NOTHING
        End Select

        Return mailSubject
    End Function

    Public Function encode(ByVal valueToEncode As String) As String
        Return HttpUtility.UrlEncode(valueToEncode)
    End Function

    Public Function decode(ByVal valueToDecode As String) As String
        Return HttpContext.Current.Server.UrlDecode(valueToDecode)
    End Function

    Function removeCharacter(ByVal stringToCleanUp, ByVal characterToRemove)
        ' replace the target with nothing
        ' Replace() returns a new String and does not modify the current one
        Return stringToCleanUp.Replace(characterToRemove, "")
    End Function

End Class
