Imports Microsoft.VisualBasic

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
            Case "PD"
                result = "Pending"
            Case "IP"
                result = "In Progress"
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

End Class
