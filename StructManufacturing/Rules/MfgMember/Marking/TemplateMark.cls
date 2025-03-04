VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberMarking
' Module: Marking Rules
'
' Description:  This class creates the BaseControl Line and Top line of Templates on Members by implementing IJDMfgSystemMarkingRule Interface
'
' Author:
'
' Comments:
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgMemberMarking.TemplateMark"
Private Const IJDMfgTemplateSet = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"


Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    Helpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    Helpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "TemplateMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
  
    On Error GoTo ErrorHandler
    
    ' Initliaze output to nothing
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = Nothing
    
    Dim oMemberPart As ISPSMemberPartPrismatic
    Dim oGlobalQuery As GSCADStructMfgGlobals.StructMfgGlobalsQuery
    Dim oTemplateSetElems As IJElements
    
    ' check if input is MemberPart
    Set oMemberPart = Part
    If oMemberPart Is Nothing Then Exit Function
    
    ' Get the TemplateSet from MemberPart
    Set oGlobalQuery = New GSCADStructMfgGlobals.StructMfgGlobalsQuery
    Set oTemplateSetElems = oGlobalQuery.GetMfgPart(oMemberPart, IJDMfgTemplateSet)
    If oTemplateSetElems Is Nothing Then GoTo CleanUp
    If oTemplateSetElems.Count < 1 Then GoTo CleanUp
    
    'Getting the Cross - Section Type of the Member
    Dim oMemberCrossSection As ISPSCrossSection
    Set oMemberCrossSection = oMemberPart.crossSection

    Dim strCrossSectionType As String
    strCrossSectionType = oMemberCrossSection.CrossSectionTypeName
    
    Dim oMGHelper As New MfgMGHelper
    
    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper

    Dim iRefObjIndex As Long
    
    Dim oTubeSurfaceBody As IJSurfaceBody
    Dim i As Integer

    If UCase(strCrossSectionType) = "PIPE" Then

        Dim lSectionID  As Long
        lSectionID = 3073  ' Outer Tube
        
        Dim oTemplateHelper         As IJMfgTemplateHelper
        Set oTemplateHelper = New MfgTemplateHelper
        
        Dim oContourElements    As IJElements
        Set oContourElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LFACES, CTX_LATERAL_LFACE, lSectionID)
        
        If oContourElements.Count > 1 Then
            Dim oSurfCollection As New Collection
            For i = 1 To oContourElements.Count
                oSurfCollection.Add oContourElements.Item(i)
            Next
            Set oTubeSurfaceBody = oMfgGeomHelper.MergeSurfaceBodies(oSurfCollection, 0.001, 0.001)
        Else
            Set oTubeSurfaceBody = oContourElements.Item(1)
        End If
        
        
        Dim oGeomCol3d As IJMfgGeomCol3d
        ' Create MfgGeomCol3d for output
        Set oGeomCol3d = m_oGeomCol3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        If oGeomCol3d Is Nothing Then GoTo CleanUp
        
        For iRefObjIndex = 1 To oTemplateSetElems.Count
            Dim oTemplateSet As IJDMfgTemplateSet
            Set oTemplateSet = oTemplateSetElems.Item(iRefObjIndex)
            
            Dim oControlLine As Object
            Dim oProjCS As IJComplexString
            
            Set oControlLine = oTemplateSet.GetControlLine
            
            Dim oCS As IJComplexString
        
            Dim oTemplateElements As IJElements
            Set oTemplateElements = oTemplateSet.GetTemplates

            Dim oTemplateBasePlane As IJPlane
            Set oTemplateBasePlane = oTemplateSet.GetBasePlane
        
            Dim oTemplate As IJMfgTemplate
            Set oTemplate = oTemplateElements.Item(1)
            
            Dim oTopCurve As IJCurve
            Set oTopCurve = oTemplate.TopLine
            
            Dim oMoniker As IMoniker
            Dim oGeom3dLine As IJMfgGeom3d
            Dim oSystemMark As IJMfgSystemMark
            Dim oMarkingInfo As IJMarkingInfo
            
            '**** Creating the Top Line of the Template as Mark ******
    
            ' The oTemplate MarkLine Elems are of type BSpCurve, hence we need to change them to ComplexString
            Set oCS = New ComplexString3d
            If Not oCS Is Nothing Then
                oCS.AddCurve oTopCurve, True
            End If
            
            Dim oProjCurve  As Object, oStartPt As Object, oEndPt As Object
            oMfgGeomHelper.IntersectSurfaceWithPlane oTubeSurfaceBody, oTemplateBasePlane, oProjCurve, oStartPt, oEndPt
            
            Dim oProjCSColl As IJElements
            oMGHelper.WireBodyToComplexStrings oProjCurve, oProjCSColl
            
            If oProjCSColl Is Nothing Then
                Set oProjCSColl = New JObjectCollection
                oProjCSColl.Add oCS
            Else
                If oProjCSColl.Count = 0 Then
                    oProjCSColl.Add oCS
                End If
            End If

            For i = 1 To oProjCSColl.Count
                Set oGeom3dLine = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                
                'Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

                oGeom3dLine.PutGeometry oProjCSColl.Item(i)
    
                oGeom3dLine.FaceId = UpSide
                oSystemMark.SetMarkingSide UpSide
                Set oMarkingInfo = oSystemMark
                oMarkingInfo.name = "Template Top Line"
                
                oGeom3dLine.PutGeometrytype STRMFG_TEMPLATE_MARK
                oSystemMark.Set3dGeometry oGeom3dLine
                
                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oTopCurve)
                oGeom3dLine.PutMoniker oMoniker
                
                oGeomCol3d.AddGeometry 1, oGeom3dLine
                
                ' ***** End of Creation of Top Line ******
                Set oGeom3dLine = Nothing
                Set oSystemMark = Nothing
                Set oMarkingInfo = Nothing
                Set oMoniker = Nothing
            Next i
            
            Set oCS = Nothing

            Set oGeom3dLine = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            
            '**** Creating the Base Control Line of the Template as Mark ******
            
            ' The oTemplateMarkLinesElems are of type BSpCurve, hence we need to change them to ComplexString
           
            Set oCS = New ComplexString3d
            oCS.AddCurve oControlLine, True
            
            'Projecting the Template Marks on the Upside or the Marking Side of the Member
            oMGHelper.ProjectComplexStringToSurface oCS, oTubeSurfaceBody, Nothing, oProjCS

            If Not oProjCS Is Nothing Then
                oGeom3dLine.PutGeometry oProjCS
            Else
                oGeom3dLine.PutGeometry oCS
            End If

            oGeom3dLine.FaceId = UpSide
            oSystemMark.SetMarkingSide UpSide
            Set oMarkingInfo = oSystemMark
            oMarkingInfo.name = "Base Control"
            
            oGeom3dLine.PutGeometrytype STRMFG_BASELINE_MARK
            
            oSystemMark.Set3dGeometry oGeom3dLine
                
            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oControlLine)
            oGeom3dLine.PutMoniker oMoniker
            
            oGeomCol3d.AddGeometry 1, oGeom3dLine
            
            ' ***** End of Creation of Base Control Line ******
            
            Set oTemplateBasePlane = Nothing
            Set oProjCurve = Nothing
            Set oStartPt = Nothing
            Set oEndPt = Nothing
            Set oProjCSColl = Nothing
            Set oTemplateElements = Nothing
            Set oTemplate = Nothing
            Set oTopCurve = Nothing
            Set oControlLine = Nothing
            Set oCS = Nothing
            Set oMarkingInfo = Nothing
            Set oSystemMark = Nothing
            Set oGeom3dLine = Nothing
            Set oMoniker = Nothing
            Set oProjCS = Nothing
        Next iRefObjIndex
    End If
    Set oTubeSurfaceBody = Nothing
    Set oMGHelper = Nothing
    Set oMfgGeomHelper = Nothing
    Set oTemplateSet = Nothing

CleanUp:
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    Set oTemplateSetElems = Nothing
    Set oMemberPart = Nothing
    Set oGlobalQuery = Nothing
    Set oGeomCol3d = Nothing
    Set oMemberCrossSection = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2017, , "RULES")

End Function
