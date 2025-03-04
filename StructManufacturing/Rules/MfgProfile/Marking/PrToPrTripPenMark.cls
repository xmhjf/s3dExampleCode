VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrToPrTripPenMark"
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

Private Const MODULE = "MfgPlateMarking.ProfPlatePenWOMark"

Private Sub Class_Initialize()
    PrMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PrMrkHelpers.UnInitialize
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
    Const METHOD = "PrToPrTripPenMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error Resume Next
        
    '*** PrToPrTripPenMark WITH Full Marks & WITHOUT Web Marks ONLY IF TRIPPING ELEMENT PRESENT ***'
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = CreatePenetrationMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, False, False, True)
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2052, , "RULES")
End Function
 




