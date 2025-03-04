VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SampleRule2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Template Default NamingRule using plate name
'
'Description:
'History :
'   ByungSam Lee              10/02/2002               Creation.
'  Marcel Veldhuizen        2004.04.22      Included correct error handling
'
'************************************************************************************************************
Option Explicit
Implements IJMfgTemplateNamingRule

Private Const MODULE = "TemplateNamingRule.SampleRule2"



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
    Dim oProcessSettings As IJMfgTemplateProcessSettings
    Dim oTemplates As IJElements, oSketchedTemplates As IJElements
    Dim oTemplate As IJMfgTemplate
    Dim oIJNamedItem As IJNamedItem
    Dim oIJNamedItemforPlate As IJNamedItem
    Dim strPlateName As String, strDirection As String
    Dim nNumTemplates As Integer, nCount As Integer
    
    Set oIJNamedItemforPlate = pPlatePart
    strPlateName = oIJNamedItemforPlate.Name
    
    Set oMfgTemplateSet = pTemplateSet
    Set oProcessSettings = oMfgTemplateSet.GetProcessSettings
    strDirection = oProcessSettings.TemplateDirection
    
    Set oTemplates = oMfgTemplateSet.GetTemplatesSorted(GirthBetweenBaseCtlPointAndAftPoint)
    nNumTemplates = oTemplates.Count
    
    'first label all templates
    For nCount = 1 To nNumTemplates
        Set oTemplate = oTemplates.Item(nCount)
        Set oIJNamedItem = oTemplate
    
        If nCount = 1 Then
            If strDirection = "Transversal" Then 'Aft Butt name
                oIJNamedItem.Name = strPlateName + "-" + "A"
            ElseIf strDirection = "Longitudinal" Or strDirection = "Waterline" Then
                oIJNamedItem.Name = strPlateName + "-" + "L"
            End If
        ElseIf nCount = nNumTemplates Then
            If strDirection = "Transversal" Then 'Fore Butt name
                oIJNamedItem.Name = strPlateName + "-" + "F"
            ElseIf strDirection = "Longitudinal" Or strDirection = "Waterline" Then
                oIJNamedItem.Name = strPlateName + "-" + "U"
            End If
        Else
            oIJNamedItem.Name = strPlateName + "-" + CStr(nCount)
        End If
        
        Set oTemplate = Nothing
        Set oIJNamedItem = Nothing
    Next nCount
    
    ' then label sketched templates
    'Sketched Template " Plate PartName + Sketched + Serial Number"
    Set oSketchedTemplates = oMfgTemplateSet.GetSketchingTemplates
 
    For nCount = 1 To oSketchedTemplates.Count
        Set oTemplate = oSketchedTemplates.Item(nCount)
        Set oIJNamedItem = oTemplate
                
        oIJNamedItem.Name = strPlateName + "-" + "Sketched" + "-" + CStr(nCount)
        
        Set oTemplate = Nothing
        Set oIJNamedItem = Nothing
    Next nCount
    
CleanUp:
    Set oMfgTemplateSet = Nothing
    Set oProcessSettings = Nothing
    Set oIJNamedItemforPlate = Nothing
    Set oTemplates = Nothing
    Set oSketchedTemplates = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3012, , "RULES")
    GoTo CleanUp
End Sub
 
