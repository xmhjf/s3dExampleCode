VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfPlatePenWOMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Profile to Plate Penetration Markings
'               WITH Full Marks WITH Web Marks
'
'  History:
'            09/24/2010      Ninad       CR-CP-184265  Ref: Asia: Mfg Ref Data + 9.13 Profile to Plate Penetration
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.ProfileToPlatePenMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function


' ***********************************************************************************
' Public Function IJDMfgSystemMarkingRule_CreateBeforeUnfold
'
' Description:  function creates fitting mark line(s) on the this plate's surface.
'               Input arguments: marking point position, this plate surface, conn plate surtface.
'               Output argument: WB of mark line.
'
'
' Unresolved issue: IsPlanar() not in plate wrapper yet
'
' ***********************************************************************************

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "ProfPlatePenWebmark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error Resume Next
        
    '*** ProfPlatePenWebmark WITH Full Marks & WITH Web Marks ***'
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = CreatePenetrationMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, True, True, False)
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1024, , "RULES")
End Function
 



