﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:2.0.50727.9164
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "9.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(ByVal sender As Global.System.Object, ByVal e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("BTMT Budgeting System")>  _
        Public ReadOnly Property ProgramTitle() As String
            Get
                Return CType(Me("ProgramTitle"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\DB\SqlCmd.xml")>  _
        Public ReadOnly Property SqlCmdFile() As String
            Get
                Return CType(Me("SqlCmdFile"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("sa")>  _
        Public ReadOnly Property Username() As String
            Get
                Return CType(Me("Username"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("BG_T_UPLOAD_DATA")>  _
        Public ReadOnly Property UploadData() As String
            Get
                Return CType(Me("UploadData"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("BG_T_ACCOUNT_REOPEN,BG_T_BUDGET_ADJUST,BG_T_BUDGET_ADJUST2,BG_T_BUDGET_DATA,BG_T_"& _ 
            "BUDGET_HEADER,BG_T_BUDGET_PERIOD,BG_T_INFORMATION")>  _
        Public ReadOnly Property BudgetData() As String
            Get
                Return CType(Me("BudgetData"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("BG_M_ACCOUNT,BG_M_ASSET_CATEGORY,BG_M_ASSET_GROUP,BG_M_ASSET_PROJECT,BG_M_BUDGET_"& _ 
            "ORDER,BG_M_CHILD_PIC,BG_M_DEPT,BG_M_PERSON_IN_CHARGE,BG_M_SETTINGS,BG_M_TRANSFER"& _ 
            "_MASTER,BG_M_USER")>  _
        Public ReadOnly Property MasterData() As String
            Get
                Return CType(Me("MasterData"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("BUDGET")>  _
        Public ReadOnly Property DatabaseName() As String
            Get
                Return CType(Me("DatabaseName"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Description : Changed MTP to MBP Migration DB Support")>  _
        Public ReadOnly Property ProgramDesc() As String
            Get
                Return CType(Me("ProgramDesc"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Data Source=10.30.99.232;Initial Catalog=BUDGET;Persist Security Info=True;User I"& _ 
            "D=sa;Password=P@ssw0rd")>  _
        Public ReadOnly Property BUDGETConnectionString() As String
            Get
                Return CType(Me("BUDGETConnectionString"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("10.30.99.232")>  _
        Public ReadOnly Property ServerName() As String
            Get
                Return CType(Me("ServerName"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("P@ssw0rd")>  _
        Public ReadOnly Property Password() As String
            Get
                Return CType(Me("Password"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("user id=sa; data source=10.30.99.232; persist security info=True; initial catalog"& _ 
            "=BUDGET;password=P@ssw0rd")>  _
        Public ReadOnly Property ConnStr() As String
            Get
                Return CType(Me("ConnStr"),String)
            End Get
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.BudgetManual.My.MySettings
            Get
                Return Global.BudgetManual.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
