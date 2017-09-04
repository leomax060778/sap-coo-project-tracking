Imports System.Data.OleDb
Imports commonLib

Partial Class Default2
    Inherits System.Web.UI.Page

    Dim sysConfiguration As New SystemConfiguration
    Dim userCommon As New SapUser
    Dim utilCommon As New Utils
    Dim appConfiguration As New AppSettings

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim actions As New SapActions
        Dim fileName1 As String = ""
        Dim fileName2 As String = ""
        Dim fileName3 As String = ""
        Dim fileName4 As String = ""
        Dim fileName5 As String = ""

        Dim deliveryButton As Button = CType(sender, Button)
        Label1.Text = ""
        If String.IsNullOrEmpty(reason.Value) Then
            Label1.Text = "Description missing. "
        Else

            If FileUpload1.HasFile Then
                Try
                    fileName1 = Now().ToString("yyyyMMdd_hhmmss_") & FileUpload1.FileName
                    FileUpload1.SaveAs(appConfiguration.deliveryStorePath & fileName1)
                    'Label1.Text = "File name: " & FileUpload1.PostedFile.FileName & "<br>" & "File Size: " & FileUpload1.PostedFile.ContentLength & " kb<br>" & "Content type: " & FileUpload1.PostedFile.ContentType

                Catch ex As System.Exception
                    Label1.Text = "ERROR: " & ex.Message.ToString()
                End Try
            Else
                Label1.Text = Label1.Text + "You have not specified a file."
            End If

            If FileUpload2.HasFile Then
                Try
                    fileName2 = Now().ToString("yyyyMMdd_hhmmss_") & FileUpload2.FileName
                    FileUpload2.SaveAs(appConfiguration.deliveryStorePath & fileName2)
                Catch ex As System.Exception
                    Label2.Text = "ERROR: " & ex.Message.ToString()
                End Try
            End If

            If FileUpload3.HasFile Then
                Try
                    fileName3 = Now().ToString("yyyyMMdd_hhmmss_") & FileUpload2.FileName
                    FileUpload3.SaveAs(appConfiguration.deliveryStorePath & fileName3)
                Catch ex As System.Exception
                    Label3.Text = "ERROR: " & ex.Message.ToString()
                End Try
            End If

            If FileUpload4.HasFile Then
                Try
                    fileName4 = Now().ToString("yyyyMMdd_hhmmss_") & FileUpload2.FileName
                    FileUpload4.SaveAs(appConfiguration.deliveryStorePath & fileName4)
                Catch ex As System.Exception
                    Label4.Text = "ERROR: " & ex.Message.ToString()
                End Try
            End If

            If FileUpload5.HasFile Then
                Try
                    fileName5 = Now().ToString("yyyyMMdd_hhmmss_") & FileUpload2.FileName
                    FileUpload5.SaveAs(appConfiguration.deliveryStorePath & fileName5)
                Catch ex As System.Exception
                    Label5.Text = "ERROR: " & ex.Message.ToString()
                End Try
            End If

            actions.deliverAI(ai_tbl_id.Text, reason.Value, fileName1, fileName2, fileName3, fileName4, fileName5)

            'Display a success message and close the form
            Button1.Enabled = False
            Response.Redirect(".\sap_owner.aspx", False)

            'Dim result As MsgBoxResult
            'result = MsgBox("AI has been delivered succesfully", MsgBoxStyle.Information, "SAP AI - Deliver")
            'If result = MsgBoxResult.Ok Then
            '    Response.Redirect(".\sap_owner.aspx", False)
            'End If

        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim dbconn As OleDbConnection
        Dim dbcomm_ais As OleDbCommand
        Dim dbread_ais As OleDbDataReader
        Dim sql_ais As String

        Dim users As New SapUser

        dbconn = New OleDbConnection(sysConfiguration.getConnection)
        dbconn.Open()

        'REQUEST ID
        Dim http_ai_id As String = Request("id")
        Dim ai_id As Integer
        Dim i As Integer

        If String.IsNullOrEmpty(http_ai_id) Or Not Integer.TryParse(http_ai_id, i) Then
            dbconn.Close()
            Response.Redirect(".\sap_owner.aspx", False)
        Else
            Integer.TryParse(http_ai_id, ai_id)
        End If

        sql_ais = "SELECT * FROM actionitems WHERE id=" + ai_id.ToString + " AND owner='" + userCommon.getId + "'"
        dbcomm_ais = New OleDbCommand(sql_ais, dbconn)

        label_ai_id.Text = ai_id.ToString

        '############ROWS#ITERATION#############
        dbread_ais = dbcomm_ais.ExecuteReader()

        If dbread_ais.HasRows Then

            dbread_ais.Read()

            Dim tRow As New HtmlTableRow

            'id
            Dim current_ai_id As Long
            current_ai_id = Convert.ToInt64(dbread_ais.GetValue(0))
            ai_tbl_id.Text = current_ai_id

            'desc
            ai_tbl_desc.Text = dbread_ais.GetString(2)
            If Len(ai_tbl_desc.Text) > 200 Then
                ai_tbl_desc.Text = Mid(ai_tbl_desc.Text, 1, 200) & "..."
            End If

            'created
            Dim ai_created As Date
            ai_created = dbread_ais.GetDateTime(3)
            ai_tbl_created.Text = ai_created.Date.ToString("dd/MMM/yyyy")

            'due
            Dim ai_duedate As Date
            If dbread_ais.IsDBNull(4) Then
                ai_tbl_due.Text = "<b>Unset</b>"
            Else
                ai_duedate = dbread_ais.GetDateTime(4)
                ai_tbl_due.Text = ai_duedate.Date.ToString("dd/MMM/yyyy")
            End If

            'desc
            ai_tbl_status.Text = utilCommon.ai_Str_Status(dbread_ais.GetString(5))

        Else

            'ERROR: AI DOES NOT EXIST
            dbconn.Close()
            dbread_ais.Close()
            Response.Redirect(".\sap_owner.aspx", False)

        End If

        dbread_ais.Close()
        dbconn.Close()
    End Sub

End Class
