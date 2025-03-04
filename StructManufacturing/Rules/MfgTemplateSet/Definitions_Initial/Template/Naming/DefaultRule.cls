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
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Template Default NamingRule
'
'Description:
'History :
'   ByungSam Lee              10/02/2002               Creation.
'   Marcel Veldhuizen       2004.04.22      Included correct error handling
'   Ram Kandimalla          2006.06.20      removed old interfaces ex: IJMfgTemplateSupportA
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
    Dim strPositionFrames As String, strPositionEven As String, strTemplateType As String, strTemplateOrientation As String
    Dim dEdgeOffset As Double, lNumOfTemplates As Long
    
    strPositionFrames = oProcessSettings.TemplatePositionFrames
    strPositionEven = oProcessSettings.TemplatePositionEven
    strTemplateType = oProcessSettings.TemplateType
    strTemplateOrientation = oProcessSettings.TemplateOrientation
    
    If strTemplateType = "Stem/Stern" Or strTemplateType = "PerpendicularXY" Then
        bSpecialType = True
    End If
    
    Dim oFrameSystem As IHFrameSystem
    Dim oMfgFrameSys As IJDMfgFrameSystem
    
    
    Set oMfgFrameSys = oMfgTemplateSet
    Set oFrameSystem = oMfgFrameSys.FrameSysParent
     
    If strPositionFrames = "PositionFrame" Then
        strProgId = oSettingsHelper.GetProgIDFromAttr("PositionFrames")
        Dim oPositionFrames As IJMfgTemplatePositionFrameRule
        Set oPositionFrames = SP3DCreateObject(strProgId)
        Set oFramesElem = oPositionFrames.GetPositionFrame(oFrameSystem, pPlatePart, oProcessSettings, oMfgTemplateSet)
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
        
    strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
    Set oPlateSideRule = SP3DCreateObject(strProgId)
    Upside = oPlateSideRule.GetPlateUpSide(pPlatePart)
    
    Dim eThicknessSide As PlateThicknessSide
    
    If (Upside = BaseSide) Then
        bBaseSide = True
        eThicknessSide = PlateBaseSide
    ElseIf (Upside = OffsetSide) Then
        bBaseSide = False
        eThicknessSide = PlateOffsetSide
    End If
    
    Dim esurfaceType As eStrMfgSurfaceType
    
    If oMfgTemplateSet.SurfaceType = 0 Then 'PART_SURFACE
        esurfaceType = TRUE_PART
    Else
        esurfaceType = TRUE_MOLD
    End If
    
    Dim nCount As Integer
    
    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    Dim strAftButt As String, strForeButt As String, strLowerSeam As String, strUpperSeam As String

    Dim oTemplates As IJElements
    Dim oTemplate As IJMfgTemplate
    Dim oIJNamedItem As IJNamedItem
    Dim nFrameIndex As Integer
    nFrameIndex = 1
    
    If strTemplateType = "Box" Or strTemplateType = "UserDefined Box" Or _
        strTemplateType = "UserDefined Box With Edges" Then
        
        Dim lGroupsCount As Long
        Dim oTemplateSetReport As IJMfgTemplateSetReport
        
        Set oTemplateSetReport = oMfgTemplateSet
        lGroupsCount = oTemplateSetReport.GroupsCount
        
        Dim oVector As New DVector
        
        'based on the direction, get the character
        Dim cDir As String
        If strDirection = "Transversal" Then
            oVector.Set 1, 0, 0
            cDir = "T"
        ElseIf strDirection = "Longitudinal" Then
            oVector.Set 0, 1, 0
            cDir = "L"
        Else
            oVector.Set 0, 0, 1
            cDir = "W"
        End If
        
        'Get Primary Direction templates
         Set oTemplates = oTemplateSetReport.GetTemplatesInGroup(1)
            
        Dim i As Long, j As Long
        For i = 1 To oTemplates.Count
            Set oTemplate = oTemplates.Item(i)
            Set oIJNamedItem = oTemplate
            oIJNamedItem.Name = cDir + CStr(i)
        Next
                
        'get the character based on secondary direction
        Dim oSurfaceBody As Object
        Set oSurfaceBody = oMfgGeomHelper.GetSurfaceFromPlate(pPlatePart, esurfaceType, eThicknessSide, 0)
        
        Dim oParallelAxisVec As IJDVector
        Dim oSecondaryDirVec As IJDVector
        Set oParallelAxisVec = GetParallelAxis(oSurfaceBody)
        Set oSecondaryDirVec = oParallelAxisVec.Cross(oVector)
           
        If Abs(oSecondaryDirVec.x) = 1 And Abs(oSecondaryDirVec.y) = 0 And Abs(oSecondaryDirVec.z) = 0 Then
            cDir = "T"
        ElseIf Abs(oSecondaryDirVec.x) = 0 And Abs(oSecondaryDirVec.y) = 1 And Abs(oSecondaryDirVec.z) = 0 Then
            cDir = "L"
        Else
            cDir = "W"
        End If
        
        Set oParallelAxisVec = Nothing
        Set oSecondaryDirVec = Nothing
        Set oSurfaceBody = Nothing
        
        'Get the secondary Direction Templates
        Set oTemplates = oTemplateSetReport.GetTemplatesInGroup(2)
        
        For i = 1 To oTemplates.Count
            Set oTemplate = oTemplates.Item(i)
            Set oIJNamedItem = oTemplate
            oIJNamedItem.Name = cDir + CStr(i)
        Next
        
        'Name the Other Templates
        For i = 3 To lGroupsCount
            Set oTemplates = oTemplateSetReport.GetTemplatesInGroup(i)
            For j = 1 To oTemplates.Count
                Set oTemplate = oTemplates.Item(j)
                Set oIJNamedItem = oTemplate
                oIJNamedItem.Name = "Group" + CStr(i) + "_" + CStr(j)
            Next
        Next
                      
    Else
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
            
            If nCount = 1 Then
                If strDirection = "Transversal" Then 'ForeButt name
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
            ElseIf nCount = nNumTemplates Then
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
            Else
                Dim oFrame As Object
                Dim strFrameName As String
                Dim oIJNamedItemForFrame As IJNamedItem
                
                If strPositionFrames = "PositionFrame" And bSpecialType = False Then
                    oIJNamedItem.Name = "T" + CStr(nCount - 1)
                    
                    Set oFrame = Nothing
                    Set oIJNamedItemForFrame = Nothing
                End If
                    
                If strPositionEven = "PositionEven" Or bSpecialType = True Then
                    Set oFrame = GetClosestAndLowestFrame(pPlatePart, oTemplate, pTemplateSet)
                    If Not oFrame Is Nothing Then
                        Set oIJNamedItemForFrame = oFrame
                        strFrameName = oIJNamedItemForFrame.Name
                        
                        oIJNamedItem.Name = strFrameName + "-" + CStr(nFrameIndex)
                    Else
                        oIJNamedItem.Name = "Template" + "-" + CStr(nFrameIndex)
                    End If
                    Set oFrame = Nothing
                    Set oIJNamedItemForFrame = Nothing
                    nFrameIndex = nFrameIndex + 1
                End If
            End If
        
            Set oTemplate = Nothing
            Set oIJNamedItem = Nothing
            
        Next nCount
        
        'Sketched Template " Plate PartName + Sketched + Serial Number"
        Dim oIJNamedItemforPlate As IJNamedItem
        Set oIJNamedItemforPlate = pPlatePart
        Dim strPlateName As String
        strPlateName = oIJNamedItemforPlate.Name
            
        For nCount = 1 To oSketchedTemplates.Count
            Set oTemplate = oSketchedTemplates.Item(nCount)
            Set oIJNamedItem = oTemplate
            
            oIJNamedItem.Name = strPlateName + "-" + "Sketched" + "-" + CStr(nCount)
        
            Set oTemplate = Nothing
            Set oIJNamedItem = Nothing
        Next nCount
    End If
    
CleanUp:
    Set oMfgTemplateSet = Nothing
    Set oProcessSettings = Nothing
    Set oSettingsHelper = Nothing
    Set oFramesElem = Nothing
    Set oTemplates = Nothing
    Set oSketchedTemplates = Nothing
    Set oMfgGeomHelper = Nothing
    Set oPlateSideRule = Nothing
    Set oIJNamedItemforPlate = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3012, , "RULES")
    GoTo CleanUp
End Sub
 
