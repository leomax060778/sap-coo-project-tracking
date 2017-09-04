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

    Public Function getSystemAdminMail() As String
        'Return "lhildt@folderit.net"

        Return "a.eyzaguirre@sap.com"
    End Function

    Public Function getSystemUrl() As String
        'result = "http://rtm-bmo.bue.sap.corp:8888/"
        ' result = "http://localhost:3542/"

        Return appConfiguration.getSystemUrl
    End Function

End Class
