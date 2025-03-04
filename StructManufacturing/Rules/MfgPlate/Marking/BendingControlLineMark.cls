VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BendingControlLineMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Bending Control Line Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH        feb. 6. 2002    created
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.BendingControlLineMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
'This rule is not yet written
Exit Function

    Const METHOD = "BendingControlMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler

    'Prepare the output collection of marking line's geometries
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_BENDING_CONTROLLINES_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
        
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part
    
    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgPart As IJMfgPlatePart
    
    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    'retrieve collection of lines
    Dim o3dCol As IJMfgGeomCol3d
    If Not oMfgPlateWrapper.GetRollInfo(o3dCol) Then
        'Since there are no roll lines we can leave the function
        GoTo CleanUp
    End If
    
    Dim iIndex As Integer
    'Iterate through the collection of lines
    For iIndex = 1 To o3dCol.Getcount
    
        'check for the type of geom (STRMFG_BENDING_CONTROLLINES_MARK)
        Dim oMfgGeom3d As IJMfgGeom3D
        Set oMfgGeom3d = o3dCol.GetGeometry(iIndex)
        If Not oMfgGeom3d.GetGeometryType = STRMFG_BENDING_CONTROLLINES_MARK Then GoTo NextLine
        
        o3dCol.RemoveGeometry oMfgGeom3d
        iIndex = iIndex - 1
    
        Dim oSystemMark As IJMfgSystemMark
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oSystemMark.SetMarkingSide UpSide
        oSystemMark.Set3dGeometry oMfgGeom3d
    
        'QI for the MarkingInfo object on the SystemMark
        Dim oMarkingInfo As MarkingInfo
        Set oMarkingInfo = oSystemMark
    
        'Set marking information, if any
    '    oMarkingInfo.PrimRollDirection = get_prim_roll_direction '?IJDVector start_point -->end_point?
    '    oMarkingInfo.SecRollDirection = get_sec_roll_direction '?IJDVector end_point --> start_point?
    '    oMarkingInfo.Radius = get_radius '?Radius of roll?
    '    oMarkingInfo.FittingAngle = get_fitting_angle 'angle of roll start_point -->end_point
    
        Dim oObjSystemMark As IUnknown
        Set oObjSystemMark = oSystemMark
   
        oGeomCol3d.AddGeometry 1, oMfgGeom3d
NextLine:
        Set oMfgGeom3d = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oObjSystemMark = Nothing
    Next
    
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Exit Function
    
CleanUp:
  
    Set oMfgPlateWrapper = Nothing
    Set oPlateWrapper = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1002, , "RULES")
    GoTo CleanUp
End Function


 
