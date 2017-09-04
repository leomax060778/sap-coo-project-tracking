Imports commonLib
Imports System.Data.OleDb

Public Class emailGenerator

    Dim appConfiguration As New AppSettings
    Dim systemConfig As New SystemConfiguration
    Dim syscfg As New SysConfig
    Dim users As New SapUser
    Dim utils As New Utils

    Public Sub createEmails()

        Dim sapMails As New MailTemplate
        Dim eLink As New Linker

        Dim report As String
        Dim reportOwner As Dictionary(Of String, String)
        Dim mail_dict As New Dictionary(Of String, String)

        Dim subject As String = "Testing email template magna aliqua."
        Dim description As String = "At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti quos dolores et quas molestias excepturi sint occaecati cupiditate non provident, similique sunt in culpa qui officia deserunt mollitia animi, id est laborum et dolorum fuga. Et harum quidem rerum facilis est et expedita distinctio. Nam libero tempore, cum soluta nobis est eligendi optio cumque nihil impedit quo minus id quod maxime placeat facere possimus, omnis voluptas assumenda est, omnis dolor repellendus. Temporibus autem quibusdam et aut officiis debitis aut rerum necessitatibus saepe eveniet ut et voluptates repudiandae sint et molestiae non recusandae. Itaque earum rerum hic tenetur a sapiente delectus, ut aut reiciendis voluptatibus maiores alias consequatur aut perferendis doloribus asperiores repellat."
        Dim detail As String = "Quiere la boca exhausta vid, kiwi, piña y fugaz jamón. Fabio me exige, sin tapujos, que añada cerveza al whisky. Jovencillo emponzoñado de whisky, ¡qué figurota exhibes! La cigüeña tocaba cada vez mejor el saxofón y el búho pedía kiwi y queso. El jefe buscó el éxtasis en un imprevisto baño de whisky y gozó como un duque. Exhíbanse politiquillos zafios, con orejas kilométricas y uñas de gavilán"
        Dim toAddress As String = users.getMailById("C5246787") + ";leonardohildt@gmail.com"
        '"lhildt@folderit.net"
        Dim reqNumber As String = "666"
        Dim aiNumber As String = "999"
        Dim aiOwner As String = "Doe, John"
        Dim requestorName As String = "Requestor, Martin"
        Dim dueDate As Date = Date.Now.AddDays(3)
        Dim dueDateAsString As String = utils.formatDateToSTring(dueDate)
        Dim dueDays As Integer = 3
        Dim dayText As String = "days"


        sapMails.imapServer = appConfiguration.imapServer
        sapMails.imapPort = appConfiguration.imapPort
        sapMails.isSSL = appConfiguration.isSSL
        sapMails.hostSMTP = appConfiguration.hostSMTP
        sapMails.emailUser = appConfiguration.emailUser
        sapMails.emailPass = appConfiguration.emailPass
        sapMails.emailAddressFrom = systemConfig.getSystemEmail()

        'Send and delete all the emails
        sapMails.sendAndDelete()

        ''Admin Report
        ''".\email-templates\SAP Email O - Admin Report.html"
        'report = getDataAdminReport()
        'mail_dict.Add("mail", "AR") 'ADMIN REPORT
        'mail_dict.Add("{admin_name}", users.getNameById("C5246787"))
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{time}", "Today")
        'mail_dict.Add("{date}", DateTime.Now.ToString("MM/dd/yyyy"))
        'mail_dict.Add("{data}", report)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the Admin report")

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        ''Owner Report
        ''".\email-templates\SAP Email N - Owner Report.html"
        'reportOwner = getDataOwnerReport()
        'mail_dict.Add("mail", "OR") 'OWNER REPORT
        'mail_dict.Add("{owner_name}", users.getNameById("C5246787"))
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{time}", "this week")
        'mail_dict.Add("{date}", utils.formatDateToSTring(Now.Date))
        'mail_dict.Add("{data}", reportOwner("this week"))
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the Owner report")
        'mail_dict.Add("{subject}", "Your Open Action Items")

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        ''Action Item Due today
        ''".\email-templates\SAP Email I - Due today.html"
        'mail_dict.Add("mail", "DL") 'AI CREATED
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{ai_id}", aiNumber)
        'mail_dict.Add("{ai_owner}", aiOwner)
        'mail_dict.Add("{description}", description) 'MAIL SUBJECT / AI DESCRIPTION
        'mail_dict.Add("{duedate}", utils.formatDateToSTring(Now.Date) & " (Today)")
        'mail_dict.Add("{delivery_link}", syscfg.getSystemUrl + "sap_dlvr.aspx?id=" + eLink.enLink(aiNumber))
        'mail_dict.Add("{requestor_name}", requestorName)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
        'mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + aiNumber)
        'mail_dict.Add("{subject}", "Your AI#" & aiNumber & " is Due Today")

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        ''Reminder due in two days
        ''".\email-templates\SAP Email B - Due 2 days.html"
        'mail_dict.Add("mail", "CF") 'AI CREATED
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{ai_id}", aiNumber)
        'mail_dict.Add("{description}", description) 'MAIL SUBJECT / AI DESCRIPTION
        'mail_dict.Add("{duedate}", dueDate.ToString("dd/MMM/yyyy"))
        'mail_dict.Add("{confirm_link}", syscfg.getSystemUrl + "sap_confirm.aspx?id=" + eLink.enLink(aiNumber))
        'mail_dict.Add("{extension_link}", syscfg.getSystemUrl + "sap_ext.aspx?id=" + eLink.enLink(aiNumber))
        'mail_dict.Add("{need_information}", syscfg.getSystemUrl + "sap_ai_data.aspx?id=" + eLink.enLink(aiNumber))
        'mail_dict.Add("{ai_owner}", aiOwner)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
        'mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + aiNumber)
        'mail_dict.Add("{subject}", "AI#" + aiNumber + " Is due in " + dueDays.ToString + " " + dayText)

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        'AI created and email to AI Owner
        '".\email-templates\SAP Email C - AI Owner.html"
        'mail_dict.Add("mail", "CR") 'AI CREATED
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{ai_id}", aiNumber)
        'mail_dict.Add("{description}", "[" + subject + "] " + description) 'MAIL SUBJECT / AI DESCRIPTION
        'mail_dict.Add("{duedate}", utils.formatDateToSTring(dueDate))
        'mail_dict.Add("{accept_link}", syscfg.getSystemUrl + "sap_accept_new_due.aspx?id=" + eLink.enLink(aiNumber))
        'mail_dict.Add("{reject_link}", syscfg.getSystemUrl + "sap_reject_due.aspx?id=" + eLink.enLink(aiNumber))
        'mail_dict.Add("{extension_link}", syscfg.getSystemUrl + "sap_ext.aspx?id=" + eLink.enLink(aiNumber))
        'mail_dict.Add("{need_information}", syscfg.getSystemUrl + "sap_ai_data.aspx?id=" + eLink.enLink(aiNumber))
        'mail_dict.Add("{ai_owner}", aiOwner)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
        'mail_dict.Add("{ai_link}", syscfg.getSystemUrl + "sap_ai_view.aspx?id=" + aiNumber)
        'mail_dict.Add("{subject}", "AI#" + aiNumber + " Is due in " + dueDays.ToString + " " + dayText)

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        ''New Request to Admin
        ''".\email-templates\SAP Email H - Admin New Request.html"
        'mail_dict.Add("mail", "NR") 'NEW REQUEST
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{rq_id}", reqNumber)
        'mail_dict.Add("{rq_link}", syscfg.getSystemUrl & "sap_req.aspx?id=" & reqNumber)
        'mail_dict.Add("{description}", "<b>" + subject + "</b><br>" + description) 'MAIL SUBJECT / AI DESCRIPTION
        'mail_dict.Add("{owners}", toAddress)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto: " & users.getAdminMail & "?subject=Questions about the report")
        'mail_dict.Add("{admin_name}", users.getNameById("C5246787"))

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        'Need more information
        '".\email-templates\SAP Email A - More info.html"
        'mail_dict.Add("mail", "ND") 'ACTION ITEM MORE INFORMATION NEEDED
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{requestor_name}", requestorName)
        'mail_dict.Add("{ai_owner}", aiOwner)
        'mail_dict.Add("{req_id}", reqNumber)
        'mail_dict.Add("{name}", subject)
        'mail_dict.Add("{hl_name}", requestorName)
        'mail_dict.Add("{description}", "Asked for more information") 'MAIL SUBJECT / AI DESCRIPTION
        'mail_dict.Add("{hl_descr}", description)
        'mail_dict.Add("{duedate}", dueDateAsString)
        'mail_dict.Add("{hl_due}", dueDateAsString)
        'mail_dict.Add("{detail}", description)
        'mail_dict.Add("{reply_mail_link}", "mailto:" & toAddress & "?subject=AI#" & aiNumber & "%20-%20To%20process%20your%20request%20additional%20info%20is%20needed&body=Dear%20Admin%2C%0D%0A%0D%0AHere%20is%20the%20info%20required%3A%0D%0A%0D%0A" & subject)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{subject}", aiOwner & " is requesting information for AI#" & aiNumber)

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        ''Delivery approval 
        ''"~/email-templates/SAP Email D - Delivery Approval.html"
        'mail_dict.Add("mail", "CP") 'NEW AI CREATED
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{ai_id}", aiNumber)
        'mail_dict.Add("{owner}", aiOwner)
        'mail_dict.Add("{description}", description)
        'mail_dict.Add("{detail}", detail)
        'mail_dict.Add("{duedate}", dueDateAsString)
        'mail_dict.Add("{delivery}", utils.formatDateToSTring(Now.Date))
        'mail_dict.Add("{requestor_name}", requestorName)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
        'mail_dict.Add("{delivery_link1}", syscfg.getSystemUrl + "delivery.ashx?file=" + "C:\test.txt")
        'mail_dict.Add("{filename1}", "C:\test.txt")
        'mail_dict.Add("{subject}", "Delivery Notice | AI#" & aiNumber)

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        'Need extension
        ''".\email-templates\SAP Email E - Extension Requested.html"
        'mail_dict.Add("mail", "NE")
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{ai_id}", aiNumber)
        'mail_dict.Add("{owner}", aiOwner)
        'mail_dict.Add("{description}", description) 'MAIL SUBJECT / AI DESCRIPTION
        'mail_dict.Add("{duedate}", dueDateAsString)
        ''utils.formatDateToSTring(ai_old_due))
        'mail_dict.Add("{extension}", dueDate.AddDays(3).ToString("dd/MMM/yyyy"))
        'mail_dict.Add("{reason}", "need more time to complete the task")
        'mail_dict.Add("{accept_link}", syscfg.getSystemUrl + "sap_accept_due.aspx?id=" + eLink.enLink(aiNumber))
        'mail_dict.Add("{reject_link}", syscfg.getSystemUrl + "sap_reject_due.aspx?id=" + eLink.enLink(aiNumber))
        'mail_dict.Add("{requestor_name}", requestorName)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
        'mail_dict.Add("{subject}", "Extension Requested for AI#" & aiNumber)

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        ''Extension approved
        ''".\email-templates\SAP Email F - Extension Approved.html"
        'mail_dict.Add("mail", "EA") 'AI EXTENSION APPROVED
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{ai_id}", aiNumber)
        'mail_dict.Add("{description}", description) 'MAIL SUBJECT / AI DESCRIPTION
        'mail_dict.Add("{duedate}", dueDateAsString)
        'mail_dict.Add("{ai_owner}", aiOwner)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
        'mail_dict.Add("{subject}", "Extension for AI#" & aiNumber & " has been approved")

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        ''Extension rejected
        ''".\email-templates\SAP Email G - Extension Rejected.html"
        'mail_dict.Add("mail", "ER") 'AI EXTENSION APPROVED
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{ai_id}", aiNumber)
        'mail_dict.Add("{description}", description) 'MAIL SUBJECT / AI DESCRIPTION
        'mail_dict.Add("{duedate}", dueDateAsString)
        'mail_dict.Add("{ai_owner}", aiOwner)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
        'mail_dict.Add("{subject}", "Extension for AI#" & aiNumber & " has been rejected")
        'mail_dict.Add("{reason}", "There is no extra time for this AI")

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        ''Action Item completed
        ''"~/email-templates/SAP Email J - AI Completed.html"
        'mail_dict.Add("mail", "AC") 'AI EXTENSION REJECTED
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{ai_id}", aiNumber)
        'mail_dict.Add("{ai_owner}", aiOwner)
        'mail_dict.Add("{description}", description) 'MAIL SUBJECT / AI DESCRIPTION
        'mail_dict.Add("{requestor_name}", requestorName)
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        'AI Not completed
        '"~/email-templates/SAP Email K - AI Uncompleted.html"
        mail_dict.Add("mail", "AU") 'AI EXTENSION REJECTED
        mail_dict.Add("to", toAddress)
        mail_dict.Add("{ai_id}", aiNumber)
        mail_dict.Add("{ai_owner}", aiOwner)
        mail_dict.Add("{description}", description) 'MAIL SUBJECT / AI DESCRIPTION
        mail_dict.Add("{duedate}", dueDateAsString)
        mail_dict.Add("{reason}", "no time to work on this")
        mail_dict.Add("{requestor_name}", requestorName)
        mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
        mail_dict.Add("{subject}", "AI#" & aiNumber & " is Incomplete")

        sapMails.SendNotificationMail(mail_dict)
        mail_dict.Clear()

        ''AI Change Owner
        ''"~/email-templates/SAP Email P - AI Owner Changed.html"
        'mail_dict.Add("mail", "OC") 'AI Owner changed
        'mail_dict.Add("to", toAddress)
        'mail_dict.Add("{ai_id}", aiNumber)
        'mail_dict.Add("{ai_owner}", aiOwner)
        'mail_dict.Add("{description}", description) 'MAIL SUBJECT / AI DESCRIPTION
        'mail_dict.Add("{app_link}", syscfg.getSystemUrl)
        'mail_dict.Add("{contact_mail_link}", "mailto:" & users.getAdminMail & "?subject=Questions about the report")
        'mail_dict.Add("{subject}", "AI#" + aiNumber.ToString + " is no longer assigned to you")

        'sapMails.SendNotificationMail(mail_dict)
        'mail_dict.Clear()

        'Send and delete all the emails
        sapMails.sendAndDelete()

    End Sub

    Public Function getDataAdminReport() As String

        Dim todayList As String = ""
        Dim count As Integer
        Dim audience As String

        'Building Owner
        'Count Acceptance Pending OW
        audience = "OWNER"
        count = ais_sum_filter("ap")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif;'>" & audience & "</td> 
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
        audience = "REQUESTOR"
        count = ais_sum_filter("rq")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & audience & "</td> 
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
        audience = "ADMIN"
        count = ais_sum_filter("dw")
        todayList = todayList & "<tr>
                                    <td style='width: 80px;border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & audience & "</td> 
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

    Private Function getDataOwnerReport() As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)
        Dim syscfg As New SysConfig
        Dim users As New SapUser

        Dim todayList As String = ""
        Dim weekList As String = ""
        Dim daysDiffStr As String = ""
        Dim dueStr As String = Now.Date.AddDays(3).ToString("dd/MMM/yyyy")
        Dim aiID As Integer = 999
        Dim aiDesc As String = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aliquam rutrum, risus eget viverra sagittis, mi lacus fermentum purus, et fermentum velit nibh id ipsum. "
        Dim newLine As String

        'PD - Pending
        newLine = Environment.NewLine & "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Pending</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>New Action Item</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"

        todayList = todayList + newLine
        weekList = weekList + newLine

        'DL Delivered
        newLine = Environment.NewLine & "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Delivered</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Already delivered</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
        todayList = todayList + newLine
        weekList = weekList + newLine

        'IP In Progress
        newLine = Environment.NewLine & "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>In Progress</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>AI in Progress</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
        todayList = todayList + newLine
        weekList = weekList + newLine

        'NE Need Extension
        newLine = Environment.NewLine & "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Need Extension</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Extension required</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
        todayList = todayList + newLine
        weekList = weekList + newLine

        'OV OVerdue
        newLine = Environment.NewLine & "<tr>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Overdue</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'> <a href=" & syscfg.getSystemUrl & "sap_ai_view.aspx?id=" & aiID.ToString & ">#" & aiID.ToString & "</a></td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>Overdue " & daysDiffStr & " day</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: left; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & aiDesc & "</td>
                                                           <td style='border-radius: 0 0 0 5px;text-align: center; font-size: 12px;font-family: Arial, Helvetica, sans-serif; color: #555555;'>" & dueStr & "</td>
                                                       </tr>"
        todayList = todayList + newLine
        weekList = weekList + newLine

        result.Add("today", todayList)
        result.Add("this week", weekList)
        Return result

    End Function

    Private Function ais_sum_filter(ByVal filter As String) As Integer
        Dim syscfg As New SysConfig
        Dim actions As New SapActions
        Dim dbconn As OleDbConnection
        Dim dbcomm As OleDbCommand
        Dim dbread As OleDbDataReader

        dbconn = New OleDbConnection(systemConfig.getConnection)
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
                extra_where = "(status IN('CR', 'PD') AND ai_count > 0)"
                extra_subq = ", (SELECT count(*) FROM actionitems WHERE requests.id = actionitems.request_id AND actionitems.status = 'DL' AND delivery_date IS NOT NULL) AS ai_count"

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
                extra_where = "((status <> 'DL' OR status <> 'PD') AND due IS NOT null AND DATEDIFF(day, TODAY(), due) >= 0 AND DATEDIFF(day, TODAY(), due) <= 7)"
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

End Class
