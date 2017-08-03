Imports System.Data.OleDb

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim syscfg As New SysConfig

        Dim su As New SapUser
        Dim ro As String = su.getRole()

        Dim actions As New SapActions

        current_user.Text = su.getName()

        If ro = "OW" Then
            Response.Redirect(syscfg.getSystemUrl + "sap_main.aspx", False)
        Else

            Dim dbconn As OleDbConnection
            Dim dbcomm, dbcomm_req, dbcomm_ais As OleDbCommand
            Dim dbread_req, dbread_ais As OleDbDataReader
            Dim sql, sql_req, sql_ais As String

            '#####TODO:#CHECK#IF#DB#EXIST###########

            dbconn = New OleDbConnection(syscfg.getConnection)
            dbconn.Open()

            'REQUEST ID
            Dim http_req_id As String = Request("id")

            'CLEAR FORM
            'descr.InnerText = ""
            'duedate.Value = ""

            'REQUEST FORM
            If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

                '#####TODO:#CHECK#VALUES################
                Dim http_fullname As String = Request.Form("fullname")
                Dim http_role As String = Request.Form("role")
                Dim http_id As String = Request.Form("id")
                Dim http_email As String = Request.Form("email")

                Dim changes As String = ""
                Dim due_updated As Boolean = False
                Dim info_updated As Boolean = False

                'GET DB ACTUAL DATA
                sql_req = "SELECT * FROM users WHERE id='" + http_req_id + "'"
                dbcomm_req = New OleDbCommand(sql_req, dbconn)
                dbread_req = dbcomm_req.ExecuteReader()

                If dbread_req.HasRows Then
                    dbread_req.Read()

                    'IF DB DATA IS DIFFERENT FROM REQUEST THEN
                    If dbread_req.GetString(4) <> http_role Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "role='" + http_role + "'"
                        info_updated = True
                    End If

                    If dbread_req.GetString(1) <> http_fullname Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "name='" + http_fullname + "'"
                        info_updated = True
                    End If

                    If dbread_req.GetString(2) <> http_email Then
                        If Not String.IsNullOrEmpty(changes) Then
                            changes = changes + ", "
                        End If
                        changes = changes + "mail='" + http_email + "'"
                        info_updated = True
                    End If

                    'IF THERE ARE CHANGES TO BE DONE
                    If Not String.IsNullOrEmpty(changes) Then
                        sql_req = "UPDATE users SET " + changes + " WHERE id='" + http_req_id + "'"
                        dbcomm_req = New OleDbCommand(sql_req, dbconn)
                        dbcomm_req.ExecuteNonQuery()

                        changes = ""
                    End If
                End If

                Dim redirectTo As String
                redirectTo = syscfg.getSystemUrl + "sap_crud.aspx"
                Response.Redirect(redirectTo, False)
            End If

            'ROWS ITERATION
            sql_req = "SELECT * FROM users WHERE id='" + http_req_id + "'"

            dbcomm_req = New OleDbCommand(sql_req, dbconn)

            'ROWS ITERATION
            '#####TODO:#CHECK#IF#REQUEST#EXIST######
            dbread_req = dbcomm_req.ExecuteReader()
            If dbread_req.HasRows Then

                dbread_req.Read()

                id.Value = dbread_req.GetValue(0).ToString
                fullname.Value = dbread_req.GetValue(1).ToString
                role.Value = dbread_req.GetString(4).ToString
                email.Value = dbread_req.GetString(2).ToString

                link_del_user.HRef = syscfg.getSystemUrl + "sap_user_del.aspx?id=" + id.Value

            End If

            dbread_req.Close()
            dbconn.Close()

        End If

    End Sub

End Class
