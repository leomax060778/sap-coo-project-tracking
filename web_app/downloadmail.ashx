<%@ WebHandler Language="VB" Class="Handler" %>

Imports System
Imports System.Web
'Imports System.Web.HttpRequest
'System.Web.HttpResponse Response = System.Web.HttpContext.Current.Response

Public Class Handler : Implements IHttpHandler
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        'context.Response.ContentType = "text/plain"
        'context.Response.Write("Hello World")
        'Dim Request As HttpRequest
        context.Response.Clear()
        'context.Response.ContentType = "application/octet-stream"
        context.Response.ContentType = "application/vnd.ms-outlook"
        'context.Response.AddHeader("content-disposition", "inline; filename=" + filename)
        'context.Response.AddHeader("Content-Disposition", "attachment; filename=" + context.Request.QueryString("mail"))
        context.Response.AddHeader("Content-Disposition", "inline; filename=" + context.Request.QueryString("mail"))
        context.Response.WriteFile("\\ARBUEWEB01\mails\" + context.Request.QueryString("mail"))
        'context.Response.BinaryWrite(Buffer)
        context.Response.End()
    End Sub
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class
