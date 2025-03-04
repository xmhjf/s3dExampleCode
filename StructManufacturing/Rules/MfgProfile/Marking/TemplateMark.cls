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
' Project: MfgProfileMarking
' Module: Marking Rules
'
' Description:  This class creates the BaseControl Line and Templatelines of Templates on Profiles by implementing IJDMfgSystemMarkingRule Interface
'
' Author:
'
' Comments:
' 2012.01.27    Santosh Ram     Added a new Class Module in Profile Marking to Support Template Marks
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgProfileMarking.TemplateMark"
Private Const IJDMfgTemplateSet = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    PrMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    PrMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "TemplateMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    ' Initliaze output to nothing
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = Nothing
    
    Dim oProfilePart As IJProfilePart
    Dim oGlobalQuery As GSCADStructMfgGlobals.StructMfgGlobalsQuery
    Dim oTemplateSetElems As IJElements
    Dim oGeomCol3d As IJMfgGeomCol3d
    ' check if input is ProfilePart
    Set oProfilePart = Part
    If oProfilePart Is Nothing Then Exit Function
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oMGHelper As New MfgMGHelper
    
    ' create MfgGeomCol3d for output
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    If oGeomCol3d Is Nothing Then GoTo CleanUp
    
    ' Get the TemplateSet from ProfilePart
    Set oGlobalQuery = New GSCADStructMfgGlobals.StructMfgGlobalsQuery
    Set oTemplateSetElems = oGlobalQuery.GetMfgPart(oProfilePart, IJDMfgTemplateSet)
    If oTemplateSetElems Is Nothing Then GoTo CleanUp
    If oTemplateSetElems.Count < 1 Then GoTo CleanUp
    
    Dim iRefObjIndex As Long
    For iRefObjIndex = 1 To ReferenceObjColl.Count
        Dim oTemplateLocationMarkObj As IUnknown
        Set oTemplateLocationMarkObj = ReferenceObjColl.Item(iRefObjIndex)

        Dim oObjSystemMark As IUnknown
        Dim oMoniker As IMoniker
        
        Dim oGeom3dLine As IJMfgGeom3d
        Set oGeom3dLine = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        
        Dim oSystemMark As IJMfgSystemMark
        'Create a SystemMark object to store additional information
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        
        Dim oMarkingInfo As IJMarkingInfo
        Dim oIJNamedItem As IJNamedItem
        Dim oTemplate As IJMfgTemplate
        Dim oCS As IJComplexString
        Dim oMfgGeomChild As IJMfgGeomChild
        Dim oSurfaceBody As IJSurfaceBody
        Dim oProjCS As IJComplexString
        Dim oTemplateSet As IJDMfgTemplateSet
        
        If TypeOf oTemplateLocationMarkObj Is IJDTemplateControlLine Then
            Set oMfgGeomChild = oTemplateLocationMarkObj
            Set oTemplateSet = oMfgGeomChild.GetParent
        ElseIf TypeOf oTemplateLocationMarkObj Is IJTemplateLocationMark Then
            Set oMfgGeomChild = oTemplateLocationMarkObj
            Dim oUnknown As IUnknown
            Set oUnknown = oMfgGeomChild.GetParent
            Set oMfgGeomChild = oUnknown
            Set oTemplateSet = oMfgGeomChild.GetParent
        End If
        
        If Not oTemplateSet Is Nothing Then
            Dim oProcessSettings As IJMfgTemplateProcessSettings
            Set oProcessSettings = oTemplateSet.GetProcessSettings
            
            Dim oSettingsHelper As MfgSettingsHelper
    
            Set oSettingsHelper = oProcessSettings
            
            Dim strProgId As String
        
            strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
        
            Dim oProfileSideRule  As IJDMfgProfileUpsideRule
            Set oProfileSideRule = SP3DCreateObject(strProgId)
            
            Dim eTemplateSide          As enumPlateSide
            eTemplateSide = oProfileSideRule.GetUpside(oProfilePart)
        End If
        
        'Creating Template Marks
        If TypeOf oTemplateLocationMarkObj Is IJTemplateLocationMark Then
            Dim oTemplateLocationMark As IJTemplateLocationMark
            
            Set oTemplateLocationMark = oTemplateLocationMarkObj

            ' The oTemplateMarkLinesElems are of type BSpCurve, hence we need to change them to ComplexString
            Set oCS = New ComplexString3d
            If Not oCS Is Nothing Then
                oCS.AddCurve oTemplateLocationMark, True
            End If

            'Projecting the Template Marks on the Upside or the Marking Side of the Profile
            Set oSurfaceBody = oProfileWrapper.GetSurfacePort(eTemplateSide).Geometry
            oMGHelper.ProjectComplexStringToSurface oCS, oSurfaceBody, Nothing, oProjCS

            If Not oProjCS Is Nothing Then
                oGeom3dLine.PutGeometry oProjCS
            Else
                oGeom3dLine.PutGeometry oCS
            End If
            oGeom3dLine.FaceId = eTemplateSide

            On Error GoTo ErrorHandler
            Set oMfgGeomChild = oTemplateLocationMark
            
            Set oIJNamedItem = oMfgGeomChild.GetParent()

            'Set the marking side
            oSystemMark.SetMarkingSide eTemplateSide
            Set oMarkingInfo = oSystemMark
            oMarkingInfo.name = oIJNamedItem.name
        
            oGeom3dLine.PutGeometrytype STRMFG_TEMPLATE_MARK
            Set oObjSystemMark = oSystemMark
            
            oSystemMark.Set3dGeometry oGeom3dLine
            
            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oTemplateLocationMark)
            oGeom3dLine.PutMoniker oMoniker
            
            oGeomCol3d.AddGeometry 1, oGeom3dLine
            Set oTemplateLocationMark = Nothing

        'Creating Base Control Line mark
        ElseIf TypeOf oTemplateLocationMarkObj Is IJDTemplateControlLine Then
            Dim oBaseControlLine As IJDTemplateControlLine
            Set oBaseControlLine = oTemplateLocationMarkObj

            Set oCS = New ComplexString3d
            oCS.AddCurve oBaseControlLine, True

            'Projecting the BaseControl Line Mark on the Upside or the Marking Side of the Profile
            Set oSurfaceBody = oProfileWrapper.GetSurfacePort(eTemplateSide).Geometry
            oMGHelper.ProjectComplexStringToSurface oCS, oSurfaceBody, Nothing, oProjCS
            
            If Not oProjCS Is Nothing Then
                oGeom3dLine.PutGeometry oProjCS
            Else
                oGeom3dLine.PutGeometry oCS
            End If
            oGeom3dLine.FaceId = eTemplateSide

            On Error GoTo ErrorHandler
            
            oSystemMark.SetMarkingSide eTemplateSide
            Set oMarkingInfo = oSystemMark
            oMarkingInfo.name = "Base Control"
    
            oGeom3dLine.PutGeometrytype STRMFG_BASELINE_MARK
    
            oSystemMark.Set3dGeometry oGeom3dLine
            
            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oBaseControlLine)
            oGeom3dLine.PutMoniker oMoniker
    
            oGeomCol3d.AddGeometry 1, oGeom3dLine
            Set oBaseControlLine = Nothing
        End If
        
        Set oIJNamedItem = Nothing
        Set oTemplate = Nothing
        Set oObjSystemMark = Nothing
        Set oMoniker = Nothing
        Set oGeom3dLine = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oCS = Nothing
        Set oSurfaceBody = Nothing
        Set oTemplateLocationMarkObj = Nothing
        Set oProjCS = Nothing
        Set oProfileSideRule = Nothing
        Set oSettingsHelper = Nothing
        Set oProcessSettings = Nothing
        Set oTemplateSet = Nothing
    Next iRefObjIndex

CleanUp:
     'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    Set oTemplateSetElems = Nothing
    Set oProfilePart = Nothing
    Set oGlobalQuery = Nothing
    Set oGeomCol3d = Nothing
    Set oMGHelper = Nothing
    
   
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2017, , "RULES")
End Function

 

