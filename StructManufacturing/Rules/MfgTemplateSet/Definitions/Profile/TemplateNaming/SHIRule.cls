VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHIRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Template NamingRule for SHI ship yard
'
'Description:
'History :
'   ByungSam Lee              10/02/2002               Creation.
' Marcel Veldhuizen         2004.04.22         Included correct error handling
'************************************************************************************************************
Option Explicit
Implements IJMfgTemplateNamingRule

Private Const MODULE = "TemplateNamingRule.SHIRule"
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
    
'    Set m_oErrors = Nothing

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
        
        Dim oFramesElem As IJElements
        Dim strProgId As String
        
        Dim bSpecialType As Boolean
        bSpecialType = False
        Dim strPositionFrames As String, strPositionEven As String, strTemplateType As String
        Dim dEdgeOffset As Double, lNumOfTemplates As Long
        strPositionFrames = oProcessSettings.TemplatePositionFrames
        strPositionEven = oProcessSettings.TemplatePositionEven
        strTemplateType = oProcessSettings.TemplateType
        
        If strTemplateType = "Stem/Stern" Or strTemplateType = "PerpendicularXY" Then
            bSpecialType = True
        End If
        
        If strPositionFrames = "PositionFrame" Then
            Dim oFrameSystem As IHFrameSystem
            Dim oMfgFrameSys As IJDMfgFrameSystem
             
            Set oMfgFrameSys = oMfgTemplateSet
            Set oFrameSystem = oMfgFrameSys.FrameSysParent
            
            strProgId = oSettingsHelper.GetProgIDFromAttr("PositionFrames")
            Dim oPositionFrames As IJMfgTemplatePositionFrameRule
            Set oPositionFrames = SP3DCreateObject(strProgId)
            Set oFramesElem = oPositionFrames.GetPositionFrame(oFrameSystem, pPlatePart, oProcessSettings)
        End If
        If strPositionEven = "PositionEven" Then
            strProgId = oSettingsHelper.GetProgIDFromAttr("PositionEven")
            Dim oPositionEven As IJMfgTemplatePositionEvenRule
            Set oPositionEven = SP3DCreateObject(strProgId)
            oPositionEven.GetPositionEven pPlatePart, dEdgeOffset, lNumOfTemplates
        End If
               
        Dim strDirection As String, strTemplateSide As String, bBaseSide As Boolean
        strDirection = oProcessSettings.TemplateDirection
        strTemplateSide = oProcessSettings.TemplateSide
        
        Dim oPlateSideRule   As IJDMfgPlateUpSideRule
        Dim Upside As enumPlateSide
                
        Set oSettingsHelper = oProcessSettings
        strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
        Set oPlateSideRule = SP3DCreateObject(strProgId)
        Upside = oPlateSideRule.GetPlateUpSide(pPlatePart)
    
        If (Upside = BaseSide) Then
            bBaseSide = True
        ElseIf (Upside = OffsetSide) Then
            bBaseSide = False
        End If
        Set oPlateSideRule = Nothing
        
        Dim nNumTemplates As Integer, nNumSketchedTemplates As Integer
        Dim oTemplates As IJElements, oSketchedTemplates As IJElements
        Set oTemplates = oMfgTemplateSet.GetTemplates
        Set oSketchedTemplates = oMfgTemplateSet.GetSketchingTemplates
        
        nNumTemplates = oTemplates.Count
        nNumSketchedTemplates = oSketchedTemplates.Count
        Dim nCount As Integer
        
        Dim oMfgGeomHelper As MfgGeomHelper
        Set oMfgGeomHelper = New MfgGeomHelper
        Dim oCS1 As IJComplexString, oCS2 As IJComplexString, oCS3 As IJComplexString, oCS4 As IJComplexString
        Dim strAftButt As String, strForeButt As String, strLowerSeam As String, strUpperSeam As String
                        
        On Error Resume Next
        oMfgGeomHelper.GetButtLinesAndSeamLines pPlatePart, bBaseSide, oCS1, oCS2, oCS3, oCS4, _
                strAftButt, strForeButt, strLowerSeam, strUpperSeam
                
        On Error GoTo ErrorHandler
        Dim oTemplate As IJMfgTemplate
        Dim oIJNamedItem As IJNamedItem
        Dim nFrameIndex As Integer
        nFrameIndex = 1
        
        For nCount = 1 To nNumTemplates
            Set oTemplate = oMfgTemplateSet.GetTemplateAtGivenIndex(nCount)
            Set oIJNamedItem = oTemplate

                If nCount = 1 Then
                    If strDirection = "Transversal" Then 'Aft Butt name
                        If strPositionEven = "PositionEven" Then
                            If strAftButt = "" Then
                                 oIJNamedItem.Name = "AftButt" + "-" + CStr(dEdgeOffset)
                            Else
                                 oIJNamedItem.Name = strAftButt + "-" + CStr(dEdgeOffset)
                            End If
                        Else
                            If strAftButt = "" Then
                                 oIJNamedItem.Name = "AftButt"
                            Else
                                 oIJNamedItem.Name = strAftButt
                            End If
                        End If
                    ElseIf strDirection = "Longitudinal" Or strDirection = "Waterline" Then
                        If strPositionEven = "PositionEven" Then
                            If strLowerSeam = "" Then
                                oIJNamedItem.Name = "LowerSeam" + "+" + CStr(dEdgeOffset)
                            Else
                                oIJNamedItem.Name = strLowerSeam + "+" + CStr(dEdgeOffset)
                            End If
                        Else
                            If strLowerSeam = "" Then
                                oIJNamedItem.Name = "LowerSeam"
                            Else
                                oIJNamedItem.Name = strLowerSeam
                            End If
                        End If
                    End If
                
            ElseIf nCount > 1 And nCount < nNumTemplates - nNumSketchedTemplates Then
                Dim oFrame As Object
                Dim strFrameName As String
                Dim oIJNamedItemForFrame As IJNamedItem
                    
                If strPositionFrames = "PositionFrame" And bSpecialType = False Then
                    Set oFrame = oFramesElem.Item(nCount)
                    Set oIJNamedItemForFrame = oFrame
                    strFrameName = oIJNamedItemForFrame.Name
                    oIJNamedItem.Name = strFrameName
                    
                    Set oFrame = Nothing
                    Set oIJNamedItemForFrame = Nothing
                End If
                
                If strPositionEven = "PositionEven" Or bSpecialType = True Then
                    Set oFrame = GetClosestAndLowestFrame(pPlatePart, oTemplate, pTemplateSet)
                    Set oIJNamedItemForFrame = oFrame
                    strFrameName = oIJNamedItemForFrame.Name
                    
                    oIJNamedItem.Name = strFrameName + "-" + CStr(nFrameIndex)
                    Set oFrame = Nothing
                    Set oIJNamedItemForFrame = Nothing
                    nFrameIndex = nFrameIndex + 1
                End If
            
            ElseIf nCount = nNumTemplates - nNumSketchedTemplates Then
                If strDirection = "Transversal" Then 'Fore Butt name
                    If strPositionEven = "PositionEven" Then
                        If strForeButt = "" Then
                            oIJNamedItem.Name = "ForeButt" + "-" + CStr(dEdgeOffset)
                        Else
                            oIJNamedItem.Name = strForeButt + "-" + CStr(dEdgeOffset)
                        End If
                    Else
                        If strForeButt = "" Then
                            oIJNamedItem.Name = "ForeButt"
                        Else
                            oIJNamedItem.Name = strForeButt
                        End If
                    End If
                ElseIf strDirection = "Longitudinal" Or strDirection = "Waterline" Then
                    If strPositionEven = "PositionEven" Then
                        If strUpperSeam = "" Then
                            oIJNamedItem.Name = "UpperSeam" + "-" + CStr(dEdgeOffset)
                        Else
                            oIJNamedItem.Name = strUpperSeam + "-" + CStr(dEdgeOffset)
                        End If
                    Else
                        If strUpperSeam = "" Then
                            oIJNamedItem.Name = "UpperSeam"
                        Else
                            oIJNamedItem.Name = strUpperSeam
                        End If
                    End If
                End If
            Else 'Sketched Template " Plate PartName + Sketched + Serial Number"
                Dim oIJNamedItemforPlate As IJNamedItem
                Set oIJNamedItemforPlate = pPlatePart
                Dim strPlateName As String
                strPlateName = oIJNamedItemforPlate.Name
                
                Dim nIndex As Integer
                nIndex = nCount - (nNumTemplates - nNumSketchedTemplates)
                oIJNamedItem.Name = strPlateName + "-" + "Sketched" + "-" + CStr(nIndex)
                
                Set oIJNamedItemforPlate = Nothing
            
            End If
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
 
