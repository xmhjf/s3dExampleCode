VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Default Rule Naming Rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit
Implements IJMfgTemplateNamingRule

Private Const MODULE = "TemplateNamingRule.DefaultRule"

Private Sub Class_Initialize()
   Const METHOD = "Class_Initialize"
   On Error GoTo ErrorHandler
   
Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJMfgTemplateNamingRule_CreateTemplateName(ByVal pProfilePart As Object, ByVal pTemplateSet As Object)
    Const METHOD = "IJMfgTemplateNamingRule_CreateTemplateName"
    On Error GoTo ErrorHandler
    
    Dim oMfgTemplateSet As IJDMfgTemplateSet
    Set oMfgTemplateSet = pTemplateSet
     
    Dim oProcessSettings As IJMfgTemplateProcessSettings
    Set oProcessSettings = oMfgTemplateSet.GetProcessSettings
    
    Dim oSettingsHelper As MfgSettingsHelper
    Set oSettingsHelper = oProcessSettings
    
    Dim oFramesElem As IJElements
    Dim strProgId As String
    
    Dim bSpecialType As Boolean
    bSpecialType = False
    Dim strPosition As String, strTemplateType As String, strTemplateOrientation As String
    Dim dEdgeOffset As Double, lNumOfTemplates As Long
    
    strPosition = oProcessSettings.TemplatePosition
    strTemplateType = oProcessSettings.TemplateType
    strTemplateOrientation = oProcessSettings.TemplateOrientation
    
    If strTemplateType = "Stem/Stern" Or strTemplateType = "PerpendicularXY" Then
        bSpecialType = True
    End If
    
    Dim oFrameSystem As IHFrameSystem
    Dim oMfgFrameSys As IJDMfgFrameSystem
    
    Set oMfgFrameSys = oMfgTemplateSet
    Set oFrameSystem = oMfgFrameSys.FrameSysParent
         
    Dim bSupportEdges As Boolean
    Dim oPosition As IJMfgTemplatePositionRule
    strProgId = oSettingsHelper.GetProgIDFromAttr("Position")
    
    If strProgId = "Ignore" Then
        'do not call the rule
    Else
        Set oPosition = SP3DCreateObject(strProgId)
        oPosition.GetPositionData oFrameSystem, pProfilePart, oProcessSettings, pTemplateSet, dEdgeOffset, lNumOfTemplates, bSupportEdges, oFramesElem
    End If
    
    Dim strDirection As String, strTemplateSide As String, bBaseSide As Boolean
    strDirection = oProcessSettings.TemplateDirection
    strTemplateSide = oProcessSettings.TemplateSide
    
    
    'TemplateSide
    strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
    
    Dim oProfileSideRule   As IJDMfgProfileUpsideRule
    Dim Upside As enumPlateSide
    Set oProfileSideRule = SP3DCreateObject(strProgId)
    Upside = oProfileSideRule.GetUpside(pProfilePart)
    
    Dim oUpsideSurface As IUnknown
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = pProfilePart
    
    Dim oSurfacePort As IJPort
    Set oSurfacePort = oProfileWrapper.GetSurfacePort(Upside)
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = oSurfacePort.Geometry
    
    Dim nCount As Integer
    
    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    Dim strAftButt As String, strForeButt As String, strLowerSeam As String, strUpperSeam As String

    Dim oTemplates As IJElements
    Dim oTemplate As IJMfgTemplate
    Dim oIJNamedItem As IJNamedItem
    Dim nFrameIndex As Integer
    nFrameIndex = 1
    
    Dim nNumTemplates As Integer, nNumSketchedTemplates As Integer
    Dim oSketchedTemplates As IJElements
    Set oTemplates = oMfgTemplateSet.GetTemplatesSorted(GirthBetweenBaseCtlPointAndForePoint)
    Set oSketchedTemplates = oMfgTemplateSet.GetSketchingTemplates
    
    nNumTemplates = oTemplates.Count
    nNumSketchedTemplates = oSketchedTemplates.Count
    
    Dim oTemplateReport As IJMfgTemplateReport
    Set oTemplateReport = oTemplates.Item(1)
    
    ' Get the seam names based on the Aft/Fore and Upper/Lower directions defined on template set.
    On Error Resume Next
    strAftButt = oTemplateReport.GetSeam(TemplateAftSeam)
    strForeButt = oTemplateReport.GetSeam(TemplateForwardSeam)
    strLowerSeam = oTemplateReport.GetSeam(TemplateLowerSeam)
    strUpperSeam = oTemplateReport.GetSeam(TemplateUpperSeam)
    On Error GoTo ErrorHandler

    For nCount = 1 To nNumTemplates
        Set oTemplate = oTemplates.Item(nCount)
        Set oIJNamedItem = oTemplate
                    
        oIJNamedItem.Name = "T" + CStr(nCount)
                    
        Set oTemplate = Nothing
        Set oIJNamedItem = Nothing
        
    Next nCount
        
    
CleanUp:
    Set oMfgTemplateSet = Nothing
    Set oProcessSettings = Nothing
    Set oSettingsHelper = Nothing
    Set oFramesElem = Nothing
    Set oTemplates = Nothing
    Set oSketchedTemplates = Nothing
    Set oMfgGeomHelper = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3012, , "RULES")
    GoTo CleanUp
End Sub
 
