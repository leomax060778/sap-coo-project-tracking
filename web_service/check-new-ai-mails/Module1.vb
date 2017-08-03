Imports System.IO
Imports commonLib

'Highligth COLOR
'#FFFC6E

Module Module1

    Dim systemConfig As New SystemConfiguration
    Dim appConfiguration As New AppSettings

    Sub MyHandler(ByVal sender As Object, ByVal args As UnhandledExceptionEventArgs)
        Dim e As Exception = DirectCast(args.ExceptionObject, Exception)
        log("Unhandled Exception: " + e.Message)
    End Sub

    Sub Main2()
        Dim sapMails As New MailTemplate
        Dim log_dict As New Dictionary(Of String, String)
        Dim newLog As New Logging
        Dim syscfg As New SysConfig
        Dim actions As New SapActions
        Dim currentDomain As AppDomain = AppDomain.CurrentDomain
        AddHandler currentDomain.UnhandledException, AddressOf MyHandler
        Dim imapClient As New MailKit.Net.Imap.ImapClient

        log("step #1: start - checking connection")

        'Testing
        'sapMails.currentEnv = "testing"
        'sapMails.imapServer = "mail.folderit.net"
        'sapMails.imapPort = 143
        'sapMails.isSSL = False
        'sapMails.hostSMTP = "mail.folderit.net"
        'sapMails.emailUser = "support_planningtool@folderit.net"
        'sapMails.emailPass = "support2016"

        'Production
        'sapMails.currentEnv = "production"
        'sapMails.imapServer = "imap.global.corp.sap"
        'sapMails.imapPort = 993
        'sapMails.isSSL = True
        'sapMails.hostSMTP = "mail.sap.corp"
        'sapMails.emailUser = "asa1_sap_mktg_in_ac@global.corp.sap\sap_marketing_in_action"
        'sapMails.emailPass = "BAoR}:qKQSkzSBO'#4pQ"
        'sapMails.emailAddressFrom = "sap_marketing_in_action@sap.com"

        sapMails.imapServer = appConfiguration.imapServer
        sapMails.imapPort = appConfiguration.imapPort
        sapMails.isSSL = appConfiguration.isSSL
        sapMails.hostSMTP = appConfiguration.hostSMTP
        sapMails.emailUser = appConfiguration.emailUser
        sapMails.emailPass = appConfiguration.emailPass
        sapMails.emailAddressFrom = systemConfig.getSystemEmail()

        Try

            log("connection in use: " + systemConfig.getConnection())

            log("step #2: Checking emails")

            'Only for testing
            'actions.sendOwnerReport("this week")
            'actions.sendOwnerReport("today")
            'actions.sendAdminReport()
            'End testing

            sapMails.CheckMail()

            log("step #3: Checking report schedule")
            log("step #3: " + String.Concat(Now.Hour.ToString(), ":", Now.Minute.ToString()))

            If Now.Hour >= 8 And Not syscfg.hasSentNotificationToday() Then
                sapMails.SendNotifications()
                syscfg.setSentNotificationToday()
            End If

            log("step #3-1: Checking Monday Report")

            'Monday - Send Weekly Owner Report
            If Now.Hour >= 8 And Today.DayOfWeek = 1 And Not syscfg.hasSentReportToday() Then
                actions.sendOwnerReport("this week")
            End If

            log("step #3-2: Checking Daily Admin Report")

            'Daily Report from Monday to Friday
            'Today.DayOfWeek <> 6 And Today.DayOfWeek <> 0 And Not syscfg.hasSentReportToday() Then
            If Now.Hour >= 8 And Not syscfg.hasSentReportToday() Then
                log("step #3-2-1: Sending Daily Admin Report")
                'actions.sendOwnerReport("today")
                actions.sendAdminReport()
                syscfg.setSentReportToday()
            End If

            log("step #3-3: Finish Daily Admin Report")

            log("step #4: Checking Send emails process")

            If syscfg.getSendMailStatus Then
                log("step #4-1: Sending emails")
                sapMails.sendAndDelete()
                log("step #4-2: Emails sent and deleted")
            End If

            Threading.Thread.Sleep(10)

            log("Email process finished")

        Catch ex As Exception
            log("Error: " & ex.Message)
        End Try

        Environment.Exit(0)

    End Sub

    Sub log(ByVal message As String)
        Dim strFile As String = "log.txt"
        Dim fileExists As Boolean = File.Exists(strFile)
        Using sw As New StreamWriter(File.Open(strFile, FileMode.Append))
            sw.WriteLine(
                    "[" & DateTime.Now & "] " & message)
        End Using
    End Sub

    Sub Main()
        Main2()
    End Sub

End Module
