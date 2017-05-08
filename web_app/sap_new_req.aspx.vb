Imports System.Data.OleDb
Imports System
Imports System.Collections.Generic
Imports Linker
Imports LogSAPTareas
Imports MailTemplate

Partial Class sap_new_req
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim mail As New MailTemplate
        Dim newMail As Dictionary(Of String, String)
        newMail = mail.CheckMail()

    End Sub
End Class
