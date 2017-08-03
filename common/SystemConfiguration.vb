Public Class SystemConfiguration

    Dim appConfiguration As New AppSettings

    Public Function getConnection() As String
        Dim result As String = ""

        Select Case appConfiguration.environment
            Case "development"
                result = "Provider=SAOLEDB;UID=root;PWD=root;Server=sap-ais-2;DBN=sap-ais-2;ASTART=No;host=localhost:2638;"
            Case "testing"
                result = "Provider=SAOLEDB;UID=root;PWD=root;Server=sap-ais-2;DBN=sap-ais-2;ASTART=No;host=localhost:2638;"
            Case "production"
                result = "Provider=SAOLEDB;UID=root;PWD=root;Server=sap-ais-2;DBN=sap-ais-2;ASTART=No;host=arbuesql01.phl.sap.corp:2638;"
        End Select

        Return result
    End Function

    Public Function getSystemEmail() As String
        Dim result As String = ""

        Select Case appConfiguration.environment
            Case "development"
                result = appConfiguration.systemEmail
            Case "testing"
                result = appConfiguration.systemEmail
            Case "production"
                result = appConfiguration.systemEmail
        End Select

        Return result
    End Function

End Class
