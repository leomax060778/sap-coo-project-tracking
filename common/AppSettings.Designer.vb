﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On



<Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
 Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "14.0.0.0")>  _
Partial Public NotInheritable Class AppSettings
    Inherits Global.System.Configuration.ApplicationSettingsBase
    
    Private Shared defaultInstance As AppSettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New AppSettings()),AppSettings)
    
    Public Shared ReadOnly Property [Default]() As AppSettings
        Get
            Return defaultInstance
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("production")>  _
    Public ReadOnly Property environment() As String
        Get
            Return CType(Me("environment"),String)
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("imap.global.corp.sap")>  _
    Public ReadOnly Property imapServer() As String
        Get
            Return CType(Me("imapServer"),String)
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("993")>  _
    Public ReadOnly Property imapPort() As Integer
        Get
            Return CType(Me("imapPort"),Integer)
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public ReadOnly Property isSSL() As Boolean
        Get
            Return CType(Me("isSSL"),Boolean)
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("mail.sap.corp")>  _
    Public ReadOnly Property hostSMTP() As String
        Get
            Return CType(Me("hostSMTP"),String)
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("asa1_sap_mktg_in_ac@global.corp.sap\sap_marketing_in_action")>  _
    Public ReadOnly Property emailUser() As String
        Get
            Return CType(Me("emailUser"),String)
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("BAoR}:qKQSkzSBO'#4pQ")>  _
    Public ReadOnly Property emailPass() As String
        Get
            Return CType(Me("emailPass"),String)
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("d:\webapps\test\mails\")>  _
    Public ReadOnly Property emailStorePath() As String
        Get
            Return CType(Me("emailStorePath"),String)
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("sap_marketing_in_action@sap.com")>  _
    Public ReadOnly Property systemEmail() As String
        Get
            Return CType(Me("systemEmail"),String)
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("http://rtm-bmo.bue.sap.corp:8888/")>  _
    Public ReadOnly Property getSystemUrl() As String
        Get
            Return CType(Me("getSystemUrl"),String)
        End Get
    End Property
    
    <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("d:\webapps\test\log.txt")>  _
    Public ReadOnly Property pathLogFile() As String
        Get
            Return CType(Me("pathLogFile"),String)
        End Get
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("d:\webapps\test\delivery\")>  _
    Public Property deliveryStorePath() As String
        Get
            Return CType(Me("deliveryStorePath"),String)
        End Get
        Set
            Me("deliveryStorePath") = value
        End Set
    End Property
End Class
