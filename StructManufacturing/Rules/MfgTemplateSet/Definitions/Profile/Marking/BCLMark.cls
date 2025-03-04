VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCLMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project: MfgTemplateMarking
'
'  Abstract:    Create Base control line mark on Template
'
'  History:
'      Siva        2nd September 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "MfgTemplateMarking.SightLine"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler
    
    Dim oMfgTemplateReport  As IJMfgTemplateReport
    Set oMfgTemplateReport = oMfgTemplate
    
    Dim oBCLPosition    As IJDPosition
    Set oBCLPosition = oMfgTemplateReport.GetPoint(BaseControlPoint)
    
    Dim oTemplateBottomCurve   As IJCurve
    Set oTemplateBottomCurve = oMfgTemplate.GetTemplateLocationMarkLine
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = GetMarkVector(oMfgTemplate, oTemplateBottomCurve, oBCLPosition)

    Dim oMarkCS     As IJComplexString
    Set oMarkCS = CreateMarkAtPosition(oBCLPosition, oMarkVec, TEMPLATE_BCL_MARK_LENGTH, True)
    
    Dim oGeomElems As IJElements
    Set oGeomElems = New JObjectCollection

    Dim oMfgGeomChild   As IJMfgGeomChild
    Set oMfgGeomChild = oMfgTemplate
    
    Dim oMfgTemplateSet As IJDMfgTemplateSet
    Set oMfgTemplateSet = oMfgGeomChild.GetParent
     
    Dim oBCL    As Object
    Set oBCL = oMfgTemplateSet.GetControlLine(oMfgTemplate.GroupIndex)
           
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New Helper
    
    Dim oGeomMoniker    As IMoniker
    Set oGeomMoniker = oMfgRuleHelper.GetMoniker(oBCL)
    
    Dim oMfgGeom3D As IJMfgGeom3d
    Set oMfgGeom3D = CreateGeom3D(oMarkCS, STRMFG_BASELINE_MARK, oGeomMoniker, "Base Control Mark")
    
    oGeomElems.Add oMfgGeom3D
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oGeomElems
    
    Exit Function

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3016, , "RULES")
End Function

