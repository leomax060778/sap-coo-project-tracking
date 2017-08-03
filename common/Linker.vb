Imports Microsoft.VisualBasic

Public Class Linker

    Public Function enLink(ByVal data As String) As String
        Return ((999999 - Val(data)) * 5 * 7 * 9 + 101010).ToString
    End Function

    Public Function deLink(ByVal data As String) As String
        Return (999999 - (Val(data) - 101010) / 5 / 7 / 9).ToString
    End Function

End Class
