VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OSSRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Template NamingRule for OSS ship yard
'
'Description:
'History :
'   ByungSam Lee              10/02/2002               Creation.
'  Marcel Veldhuizen        2004.04.22      Included correct error handling
'
'************************************************************************************************************
Option Explicit
Implements IJMfgTemplateNamingRule
Private Const MODULE = "TemplateNamingRule.OSSRule"
'Private Const E_FAIL = &H80004005      ' For the error object.
'Private m_oErrors As IJEditErrors      ' To collect and propagate the errors.
'Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"

Private Sub Class_Initialize()
   Const METHOD = "Class_Initialize"
   On Error GoTo ErrorHandler
   
   
'   Set m_oErrors = Nothing
'   Set m_oErrors = CreateObject(ERRORPROGID)

Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Class_Terminate()

'Set m_oErrors = Nothing

End Sub

Private Sub IJMfgTemplateNamingRule_CreateTemplateName(ByVal pPlatePart As Object, ByVal pTemplateSet As Object)
    Const METHOD = "IJMfgTemplateNamingRule_CreateTemplateName"
    On Error GoTo ErrorHandler
        Dim oMfgTemplateSet As IJDMfgTemplateSet
        Set oMfgTemplateSet = pTemplateSet
                
        Dim oProcessSettings As IJMfgTemplateProcessSettings
        Set oProcessSettings = oMfgTemplateSet.GetProcessSettings
        
        Dim oSettingsHelper As MfgSettingsHelper
        Set oSettingsHelper = oProcessSettings
        
        Dim strProgId As String
        
        Dim strPositionEven As String
        Dim dEdgeOffset As Double, lNumOfTemplates As Long
        strPositionEven = oProcessSettings.TemplatePositionEven
        
        If strPositionEven = "PositionEven" Then
            strProgId = oSettingsHelper.GetProgIDFromAttr("PositionEven")
            Dim oPositionEven As IJMfgTemplatePositionEvenRule
            Set oPositionEven = SP3DCreateObject(strProgId)
            oPositionEven.GetPositionEven pPlatePart, dEdgeOffset, lNumOfTemplates
        End If
               
        Dim strDirection As String
        strDirection = oProcessSettings.TemplateDirection
        
        Dim nNumTemplates As Integer, nNumSketchedTemplates As Integer
        Dim oTemplates As IJElements, oSketchedTemplates As IJElements
        Set oTemplates = oMfgTemplateSet.GetTemplates
        Set oSketchedTemplates = oMfgTemplateSet.GetSketchingTemplates
        
        nNumTemplates = oTemplates.Count
        nNumSketchedTemplates = oSketchedTemplates.Count
        Dim nCount As Integer
        
        On Error GoTo ErrorHandler
        Dim oTemplate As IJMfgTemplate
        Dim oIJNamedItem As IJNamedItem
        
        Dim oIJNamedItemforPlate As IJNamedItem
        Set oIJNamedItemforPlate = pPlatePart
        Dim strPlateName As String
        strPlateName = oIJNamedItemforPlate.Name
        
        For nCount = 1 To nNumTemplates
            Set oTemplate = oMfgTemplateSet.GetTemplateAtGivenIndex(nCount)
            Set oIJNamedItem = oTemplate

            If nCount = 1 Then
                If strDirection = "Transversal" Then 'Aft Butt name
                    oIJNamedItem.Name = strPlateName + "-" + "A"
                ElseIf strDirection = "Longitudinal" Or strDirection = "Waterline" Then
                    oIJNamedItem.Name = strPlateName + "-" + "L"
                End If
                
            ElseIf nCount > 1 And nCount < nNumTemplates - nNumSketchedTemplates Then
                oIJNamedItem.Name = strPlateName + "-" + CStr(nCount)
                
            ElseIf nCount = nNumTemplates - nNumSketchedTemplates Then
                If strDirection = "Transversal" Then 'Fore Butt name
                    oIJNamedItem.Name = strPlateName + "-" + "F"
                ElseIf strDirection = "Longitudinal" Or strDirection = "Waterline" Then
                    oIJNamedItem.Name = strPlateName + "-" + "U"
                End If
            Else 'Sketched Template " Plate PartName + Sketched + Serial Number"
                Dim nIndex As Integer
                nIndex = nCount - (nNumTemplates - nNumSketchedTemplates)
                oIJNamedItem.Name = strPlateName + " Sketched " + CStr(nIndex)
                
                Set oIJNamedItemforPlate = Nothing
            
            End If
            Set oTemplate = Nothing
            Set oIJNamedItem = Nothing

        Next nCount
        
CleanUp:
        Set oMfgTemplateSet = Nothing
        Set oProcessSettings = Nothing
        Set oSettingsHelper = Nothing
        Set oTemplates = Nothing
        Set oSketchedTemplates = Nothing
        
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3012, , "RULES")
    GoTo CleanUp
End Sub
 
