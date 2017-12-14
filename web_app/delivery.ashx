<%@ WebHandler Language="VB" Class="Handler" %>

Imports System
Imports System.Web

Public Class Handler : Implements IHttpHandler

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "application/octet-stream"
        context.Response.AddHeader("Content-Disposition", "attachment; filename=" + context.Request.QueryString("file"))
        context.Response.WriteFile("\\ARBUEWEB01\delivery\" + context.Request.QueryString("file"))
        'context.Response.WriteFile("\delivery\" + context.Request.QueryString("file"))
        context.Response.End()
    End Sub

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class
