VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Default"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract   :
'   Tube Template Default NamingRule
'
'Description :
'History     :
'   Siva              20th October 2011               Creation.
'************************************************************************************************************
Option Explicit
Implements IJMfgTemplateNamingRule

Private Const MODULE = "StrMfgTemplateNamingTube.Default"

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
   
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJMfgTemplateNamingRule_CreateTemplateName(ByVal pPlatePart As Object, ByVal pTemplateSet As Object)
    Const METHOD = "IJMfgTemplateNamingRule_CreateTemplateName"
    On Error GoTo ErrorHandler
    
    Dim oMfgTemplateSet As IJDMfgTemplateSet
    Set oMfgTemplateSet = pTemplateSet
    
    Dim oProcessSettings As IJMfgTemplateProcessSettings
    Set oProcessSettings = oMfgTemplateSet.GetProcessSettings
    
    Dim oSettingsHelper  As MfgSettingsHelper
    Set oSettingsHelper = oProcessSettings
    
    Dim strTemplateSideEnd As String
    Dim bstrTableName As String
   
    bstrTableName = oSettingsHelper.GetTableNameFromAttr("SideEnd")
    strTemplateSideEnd = oSettingsHelper.GetShortDescription(bstrTableName, "SideEnd")
    
    Dim oTemplates As IJElements
    Set oTemplates = oMfgTemplateSet.GetTemplates
    
    Dim oIJNamedItem As IJNamedItem
    Set oIJNamedItem = oTemplates.Item(1)
    
    oIJNamedItem.Name = strTemplateSideEnd + "-" + "Template"
    
    Set oIJNamedItem = Nothing
    
CleanUp:
    Set oMfgTemplateSet = Nothing
    Set oProcessSettings = Nothing
    Set oTemplates = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3012, , "RULES")
    GoTo CleanUp
End Sub
