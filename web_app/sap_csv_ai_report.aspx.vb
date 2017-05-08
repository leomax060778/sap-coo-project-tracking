Imports System.Data.OleDb
Imports System
Imports System.Collections.Generic
Imports Linker
Imports LogSAPTareas
Imports MailTemplate
Imports SapActions
Imports SapAnalytics

Partial Class sap_csv_ai_report
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim actions As New SapActions
        Dim anal As New SapAnalytics
        report_table.Text = anal.createActionItemsReport(New Date(2015, 1, 1), Today.Date)
    End Sub

End Class
