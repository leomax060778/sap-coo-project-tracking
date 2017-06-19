Imports Microsoft.VisualBasic

Imports System.Data.OleDb
Imports System

Public Class SapUser

    Public Function getId() As String
        Dim ru As String = System.Web.HttpContext.Current.User.Identity.Name
        If ru = "" Then
            ru = System.Environment.UserName
            '"I828136 - Nico Morales - Requestor" 
            'I821137 - Martin Whitehead - Admin"
            'I858826 - Eyzaguirre, Angiecarla
            ru = "I828136"
        Else
            ru = Mid(ru, InStr(ru, "\") + 1)
        End If
        getId = ru
    End Function

    Public Function getName() As String
        getName = getNameById(getId)
    End Function

    Public Function getNameById(ByVal id As String) As String
        Dim result As String
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT name FROM users WHERE id='" & id & "'"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.GetString(0)
        Else
            result = "guest"
        End If

        dbconn.Close()

        Return result
    End Function

    Public Function getIdByMail(ByVal mailAddress As String) As String
        Dim result As String
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT id FROM users WHERE mail='" & mailAddress & "'"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.GetString(0)
        Else
            result = "guest"
        End If

        dbconn.Close()

        Return result
    End Function

    Public Function getNameByMail(ByVal mailAddress As String) As String
        Dim result As String
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT name FROM users WHERE mail='" & mailAddress & "'"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.GetString(0)
        Else
            result = "guest"
        End If

        dbconn.Close()

        Return result
    End Function

    Public Function getMailById(ByVal data As String) As String
        Dim result As String
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT mail FROM users WHERE id='" & data & "'"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.GetString(0)
        Else
            result = "guest"
        End If

        dbconn.Close()

        Return result
    End Function

    Public Function getRole() As String
        getRole = getRoleById(getId)
    End Function


    Public Function getRoleById(ByVal id As String) As String
        Dim result As String
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT role FROM users WHERE id='" & id & "'"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            dbread_req.Read()
            result = dbread_req.GetString(0)
        Else
            result = "guest"
        End If

        dbconn.Close()

        Return result
    End Function

    Public Function getOwners(ByVal requestor As String) As List(Of String)
        Dim result As New List(Of String)
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT * FROM users WHERE role='OW'"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            While dbread_req.Read()
                result.Add(dbread_req.GetString(2))
            End While
        End If

        dbconn.Close()

        Return result
    End Function

    Public Function isRequestor(ByVal mailAddress As String) As Boolean
        Dim result As Boolean = False
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String
        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()
        sql_req = "SELECT * FROM users WHERE role='RQ' AND mail='" + mailAddress + "'"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then 'FOUND THE MAIL AS REQUESTOR
            result = True
        End If

        dbconn.Close()

        Return result
    End Function

    Public Function isOwner(ByVal mailAddress As String) As Boolean
        Dim result As Boolean = False
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String
        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()
        sql_req = "SELECT * FROM users WHERE (role='OW' OR role='AO') AND mail='" + mailAddress + "'"
        dbcomm_req = New OleDbCommand(sql_req, dbconn)
        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then 'FOUND THE MAIL AS OWNER
            result = True
        End If

        dbconn.Close()

        Return result
    End Function

    Function getAdminMail() As String
        Dim result As String = ""
        Dim dbconn As OleDbConnection
        Dim dbcomm_req As OleDbCommand
        Dim dbread_req As OleDbDataReader
        Dim sql_req As String

        Dim syscfg As New SysConfig

        dbconn = New OleDbConnection(syscfg.getConnection)
        dbconn.Open()

        sql_req = "SELECT mail FROM users WHERE role='AO' OR role='AD'"

        dbcomm_req = New OleDbCommand(sql_req, dbconn)

        dbread_req = dbcomm_req.ExecuteReader()
        If dbread_req.HasRows Then
            While dbread_req.Read()
                result = result & dbread_req.GetString(0) & ";"
            End While
        End If

        dbconn.Close()

        Return result
    End Function

End Class
