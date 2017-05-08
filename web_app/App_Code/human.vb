Imports Microsoft.VisualBasic

Public Class Humanize

    Public Function date_fwd(ByVal i As Integer) As String
        Dim result As String
        Select Case i
            Case Is < 0
                result = "Overdue"
            Case 0
                result = "Today"
            Case 1
                result = "Tomorrow"
            Case Else
                result = i.ToString
        End Select
        Return result
    End Function

    Public Function date_bkw(ByVal i As Integer) As String
        Dim result As String
        Select Case i
            Case Is < 0
                result = "Error"
            Case 0
                result = "Today"
            Case 1
                result = "Yesterday"
            Case Else
                result = i.ToString
        End Select
        Return result
    End Function

End Class
