VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'    TemplSetData object encapsulates the attributes such as side, orientation etc. of a TemplateSet.
'
'Description:
'History :
'   Siva              20th Oct 2011      Creation
'******************************************************************************************************************

Option Explicit

Private Const MODULE = "StrMfgTemplateProcessEdge.EdgeTemplSetData"

Private m_lSide As Long
Private m_dOffset As Double, m_dExtn_Base As Double, m_dExtn_Offset As Double

Public Sub InitSettings(ByVal oPart As Object, ByVal oProcessSettings As IJMfgTemplateProcessSettings, ByVal oTemplateSet As IJDMfgTemplateSet)
    Const METHOD = "InitSettings"
    On Error Resume Next
        
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oPart Is Nothing Then
        strErrorMsg = strErrorMsg & "oPart is Nothing "
        GoTo ErrorHandler
    End If
    If oProcessSettings Is Nothing Then
        strErrorMsg = strErrorMsg & "oProcessSettings is Nothing "
        GoTo ErrorHandler
    End If
    
    Dim oSettingsHelper  As MfgSettingsHelper
    Set oSettingsHelper = oProcessSettings
    
    If oSettingsHelper Is Nothing Then
        strErrorMsg = strErrorMsg & "oSettingsHelper is Nothing "
        GoTo ErrorHandler
    End If
    
    ' Get values from template process settings object
    ' TemplateType
    strErrorMsg = "TemplateType failed"
    
    Dim strProgId As String
   
    'Side
    strErrorMsg = "Side failed"
    strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
    Dim oSideRule As IJDMfgEdgeTemplateSideRule
    Set oSideRule = SP3DCreateObject(strProgId)
    m_lSide = oSideRule.GetTemplateSide(oPart)
    
    'Offset
    strErrorMsg = "Offset failed"
    strProgId = oSettingsHelper.GetProgIDFromAttr("Offset")
    Dim oOffsetRule As IJDMfgEdgeTemplateOffsetRule
    Set oOffsetRule = SP3DCreateObject(strProgId)
    m_dOffset = oOffsetRule.GetTemplateOffset(oPart, m_lSide)
        
    'Extension
    strErrorMsg = "Extension failed"
    strProgId = oSettingsHelper.GetProgIDFromAttr("Extension")
    Dim oExtensionRule As IJMfgEdgeTemplateExtensionRule
    Set oExtensionRule = SP3DCreateObject(strProgId)
    oExtensionRule.GetTemplateExtension oPart, m_dExtn_Base, m_dExtn_Offset
    
CleanUp:
    Set oSettingsHelper = Nothing
    Set oSideRule = Nothing
    Set oOffsetRule = Nothing
    Set oExtensionRule = Nothing
    
    Exit Sub
    
GSWF_ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_FailedToGetSurfaceFromPlatePart, , "RULES")
    GoTo CleanUp
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

Public Property Get TemplateSide() As Long
    TemplateSide = m_lSide
End Property

Public Property Get TemplateOffset() As String
    TemplateOffset = m_dOffset
End Property

Public Property Get TemplateBaseExtension() As Double
    TemplateBaseExtension = m_dExtn_Base
End Property

Public Property Get TemplateOffsetExtension() As Double
    TemplateOffsetExtension = m_dExtn_Offset
End Property
