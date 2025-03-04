VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseControlMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation.  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Base Control Line Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH        feb. 6. 2002    created
'      MJV  2004.04.23          Include correct error handling
'      Thakur 2005.April.25/26  Modified the implementation in CreateBeforeUnfold to support
'                               Multiple templateset for a plate.
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.BaseControlMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "BaseControlMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    ' Initliaze output to nothing
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = Nothing
    
    Dim oPlatePart As IJPlatePart
    Dim strTemplateSide As String
    Dim oObjSystemMark As IUnknown
    Dim oMoniker As IMoniker
    Dim oGeom3dLine As IJMfgGeom3D
    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As IJMarkingInfo
    Dim oCS As IJComplexString
        
    Dim oGeomCol3d As IJMfgGeomCol3d
    Dim oMfgGeom3dFactory As New MfgGeom3dFactory

    ' check if input is PlatePart
    Set oPlatePart = Part
    If oPlatePart Is Nothing Then Exit Function
    
    ' create MfgGeomCol3d for output
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    If oGeomCol3d Is Nothing Then GoTo CleanUp
    
    CreateAPSMarkings STRMFG_BASELINE_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

    ' In V6 there can be more than one TemplateSet per PlatePart
    Dim iRefObjIndex As Long
    For iRefObjIndex = 1 To ReferenceObjColl.Count
        Dim oBaseControlLineObj As IUnknown
        Set oBaseControlLineObj = ReferenceObjColl.Item(iRefObjIndex)
        
        If TypeOf oBaseControlLineObj Is IJDTemplateControlLine Then
            Dim oMfgDef As IJMfgDefinition
            Dim oConnectedPort As IJPort
            Set oMfgDef = oBaseControlLineObj
            Set oConnectedPort = oMfgDef.GetPort
            If oConnectedPort.Connectable Is Part Then           
				'**************************** Modified Part **********************************
				
				Dim oGeomChild As IJMfgGeomChild
				
				Set oGeomChild = oBaseControlLineObj
				
				Dim oTemplateSet As IJDMfgTemplateSet
				Set oTemplateSet = oGeomChild.GetParent
				
				Dim oProcessSettings As IJMfgTemplateProcessSettings
				Set oProcessSettings = oTemplateSet.GetProcessSettings
				
				Dim oSettingsHelper As MfgSettingsHelper
		
				Set oSettingsHelper = oProcessSettings
				
				Dim strProgId As String
			
				strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
			
				Dim oPlateSideRule  As IJDMfgPlateUpSideRule
				Set oPlateSideRule = SP3DCreateObject(strProgId)
				
				Dim eTemplateSide          As enumPlateSide
				eTemplateSide = oPlateSideRule.GetPlateUpSide(oPlatePart)
				
				
			   '*****************************************************************************
			   
			
				Dim oBaseControlLine As IJDTemplateControlLine
				Set oBaseControlLine = oBaseControlLineObj
				Set oGeom3dLine = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
				 
				' BaseControlLine may be a ComplexString if plate has patches
				' Query for IJComplexString on oBaseControlLine
				' Ignore errors since query may fail in some cases
				On Error Resume Next
				Set oCS = oBaseControlLine
				If oCS Is Nothing Then
					Set oCS = New ComplexString3d
					oCS.AddCurve oBaseControlLine, True
				End If
				oGeom3dLine.PutGeometry oCS
				On Error GoTo ErrorHandler
				
				'Create a SystemMark object to store additional information
				Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
				'Set the marking side
				oSystemMark.SetMarkingSide eTemplateSide
				Set oMarkingInfo = oSystemMark
				oMarkingInfo.Name = "Base Control"
		
				oGeom3dLine.PutGeometrytype STRMFG_BASELINE_MARK
		
				oSystemMark.Set3dGeometry oGeom3dLine
				
				Set oMoniker = m_oMfgRuleHelper.GetMoniker(oBaseControlLine)
				oGeom3dLine.PutMoniker oMoniker
		
				oGeomCol3d.AddGeometry 1, oGeom3dLine
		
				Set oBaseControlLine = Nothing
				Set oMoniker = Nothing
				Set oGeom3dLine = Nothing
				Set oSystemMark = Nothing
				Set oMarkingInfo = Nothing
				Set oCS = Nothing
				Set oObjSystemMark = Nothing
			End If
        End If
    Next iRefObjIndex

    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oObjSystemMark = Nothing
    Set oMoniker = Nothing
    Set oGeom3dLine = Nothing
    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing
    Set oCS = Nothing
    Set oPlatePart = Nothing
    Set oGeomCol3d = Nothing
    Set oMfgGeom3dFactory = Nothing
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1001, , "RULES")
    GoTo CleanUp
End Function

 
