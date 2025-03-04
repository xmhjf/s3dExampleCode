VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RobotMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Robot Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH        feb. 6. 2002    created
'       KONI        june 22 2002    Tested OK
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.RobotMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "RobotMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d      As IJMfgGeomCol3d
    Dim oMarkingObjColl As JCmnShp_CollectionAlias

    Set oMarkingObjColl = GetMarkingLinesForThisPart(Part, ReferenceObjColl)
    Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(oMarkingObjColl, STRMFG_ROBOT_MARK)
    
    If Not oGeomCol3d Is Nothing Then
        Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = SetMarkinLineNameOnGeom3DObjects(oGeomCol3d)
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1025, , "RULES")
End Function

 
