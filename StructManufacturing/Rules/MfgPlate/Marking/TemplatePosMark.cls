VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplatePosMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Template Position Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH        feb. 6. 2002    created
'      MJV        2004.04.23      Included correct error handling
'    Ram Kandimalla 2005.07.01    Changed CreateBeforeUnfold to make use of methods from TemplateSet
'                                 and added support for multiple templates
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.TemplatePosMark"
Private Const IJDMfgTemplateSet = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"
Private m_SupportPlatePart As IJMfgWrapPlatePart ' The support class.


Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "TemplatePosMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler

    ' Initliaze output to nothing
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = Nothing
    
    Dim oPlatePart As IJPlatePart
    Dim oGlobalQuery As GSCADStructMfgGlobals.StructMfgGlobalsQuery
    Dim oTemplateSetElems As IJElements
    Dim oGeomCol3d As IJMfgGeomCol3d
    ' check if input is PlatePart
    Set oPlatePart = Part
    If oPlatePart Is Nothing Then Exit Function
    
    ' create MfgGeomCol3d for output
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    If oGeomCol3d Is Nothing Then GoTo CleanUp
    
    CreateAPSMarkings STRMFG_TEMPLATE_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    ' Get the TemplateSet from PlatePart
    Set oGlobalQuery = New GSCADStructMfgGlobals.StructMfgGlobalsQuery
    Set oTemplateSetElems = oGlobalQuery.GetMfgPart(oPlatePart, IJDMfgTemplateSet)
    If oTemplateSetElems Is Nothing Then GoTo CleanUp
    If oTemplateSetElems.Count < 1 Then GoTo CleanUp
        
    ' In V6 there can be more than one TemplateSet per PlatePart
    Dim iRefObjIndex As Long
    For iRefObjIndex = 1 To ReferenceObjColl.Count
        Dim oTemplateLocationMarkObj As IUnknown
        Set oTemplateLocationMarkObj = ReferenceObjColl.Item(iRefObjIndex)
        
        If TypeOf oTemplateLocationMarkObj Is IJTemplateLocationMark Then
            Dim oMfgDef As IJMfgDefinition
            Dim oConnectedPort As IJPort
            Set oMfgDef = oTemplateLocationMarkObj
            Set oConnectedPort = oMfgDef.GetPort
            If oConnectedPort.Connectable Is Part Then		
				Dim oObjSystemMark As IUnknown
				Dim oMoniker As IMoniker
				Dim oGeom3dLine As IJMfgGeom3D
				Dim oSystemMark As IJMfgSystemMark
				Dim oMarkingInfo As IJMarkingInfo
				Dim oIJNamedItem As IJNamedItem
				Dim oTemplate As IJMfgTemplate
				Dim oCS As IJComplexString
				Dim oTemplateLocationMark As IJTemplateLocationMark
				
				Set oTemplateLocationMark = oTemplateLocationMarkObj
				
				Set oGeom3dLine = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
				On Error Resume Next
				' TR-CP·94841  Template set marking lines are not created for mfg part
				' On some plates the oTemplateMarkLinesElems are of type IJComplexString
				Set oCS = oTemplateLocationMark
				If oCS Is Nothing Then
					' The oTemplateMarkLinesElems are of type BSpCurve, hence we need to change them to ComplexString
					Set oCS = New ComplexString3d
					If Not oCS Is Nothing Then
						oCS.AddCurve oTemplateLocationMark, True
					End If
				End If
				oGeom3dLine.PutGeometry oCS
				On Error GoTo ErrorHandler
					
				Dim oMfgGeomChild As IJMfgGeomChild
				Set oMfgGeomChild = oTemplateLocationMark
				
				Set oIJNamedItem = oMfgGeomChild.GetParent()
		
				'Create a SystemMark object to store additional information
				Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
				'Set the marking side
				oSystemMark.SetMarkingSide UpSide
				Set oMarkingInfo = oSystemMark
				oMarkingInfo.Name = oIJNamedItem.Name
			
				oGeom3dLine.PutGeometrytype STRMFG_TEMPLATE_MARK
				Set oObjSystemMark = oSystemMark
				
				oSystemMark.Set3dGeometry oGeom3dLine
				
				Set oMoniker = m_oMfgRuleHelper.GetMoniker(oTemplateLocationMark)
				oGeom3dLine.PutMoniker oMoniker
				
				oGeomCol3d.AddGeometry 1, oGeom3dLine
		
				Set oIJNamedItem = Nothing
				Set oTemplate = Nothing
				Set oObjSystemMark = Nothing
				Set oMoniker = Nothing
				Set oGeom3dLine = Nothing
				Set oSystemMark = Nothing
				Set oMarkingInfo = Nothing
				Set oCS = Nothing
			End If
        End If
    Next iRefObjIndex
	
CleanUp:
    ' CreateElongationMarks Part, UpSide, ReferenceObjColl, oGeomCol3d
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Set oTemplateSetElems = Nothing
    Set oPlatePart = Nothing
    Set oGlobalQuery = Nothing
    Set oGeomCol3d = Nothing
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1030, , "RULES")
    GoTo CleanUp
End Function

Private Sub CreateElongationMarks(ByVal Part As Object, ByVal UpSide As Long, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal oGeomColl3d As GSCADMfgRulesDefinitions.IJMfgGeomCol3d)
    Const METHOD = "TemplatePosMark: CreateElongationMarks"
    
    Dim iRefObjIndex As Long
    Dim oTemplateMarkColl As New Collection
    
    Dim oMarkingObjColl As JCmnShp_CollectionAlias

    Set oMarkingObjColl = GetMarkingLinesForThisPart(Part, ReferenceObjColl)
    
    For iRefObjIndex = 1 To ReferenceObjColl.Count
        Dim oTemplateLocationMarkObj As IUnknown
        Set oTemplateLocationMarkObj = ReferenceObjColl.Item(iRefObjIndex)
        
        If TypeOf oTemplateLocationMarkObj Is IJTemplateLocationMark Then
            Dim oMfgDef As IJMfgDefinition
            Dim oConnectedPort As IJPort
            Set oMfgDef = oTemplateLocationMarkObj
            Set oConnectedPort = oMfgDef.GetPort
            If oConnectedPort.Connectable Is Part Then
                oTemplateMarkColl.Add oTemplateLocationMarkObj
            End If
        End If
    Next iRefObjIndex
    
    For iRefObjIndex = 1 To oMarkingObjColl.Count
        Dim oTemplateManualMarkObj As IUnknown
        Set oTemplateManualMarkObj = oMarkingObjColl.Item(iRefObjIndex)
        
        If TypeOf oTemplateManualMarkObj Is IJMfgMarkingLinesData Then
            Dim lMarkingType As Long
            Dim oMfgMarkingLinesData As IJMfgMarkingLinesData
            
            Set oMfgMarkingLinesData = oMarkingObjColl.Item(iRefObjIndex)
            lMarkingType = oMfgMarkingLinesData.GetMfgMarkingType

            If lMarkingType = STRMFG_ELONGATION_MARK Then
                oTemplateMarkColl.Add oTemplateManualMarkObj
            End If
            Set oMfgMarkingLinesData = Nothing
        End If
    Next iRefObjIndex
    
    Dim oRuleHlpr As New MfgRuleHelpers.Helper
    
    oRuleHlpr.SplitForElongationMarks Part, oTemplateMarkColl, oGeomColl3d, 0.35
    
CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1030, , "RULES")
    GoTo CleanUp
End Sub
