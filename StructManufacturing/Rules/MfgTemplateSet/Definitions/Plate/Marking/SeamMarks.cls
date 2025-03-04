VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project: MfgTemplateMarking
'
'  Abstract:    Creates Seam marks on Template
'
'  History:
'      Siva        2nd September 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingPlate.SeamMarks"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler
    
    Dim oMfgTemplateReport  As IJMfgTemplateReport
    Set oMfgTemplateReport = oMfgTemplate
    
    Dim oTemplateBottomCurve   As IJCurve
    Set oTemplateBottomCurve = oMfgTemplate.GetTemplateLocationMarkLine
    
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New Helper
    
    Dim oGeomElems As IJElements
    Set oGeomElems = New JObjectCollection
    
    Dim oLowerSeamPosition    As IJDPosition
    Set oLowerSeamPosition = oMfgTemplateReport.GetPoint(LowerSeamPoint)
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = GetMarkVector(oMfgTemplate, oTemplateBottomCurve, oLowerSeamPosition)

    Dim oMarkCS     As IJComplexString
    Set oMarkCS = CreateMarkAtPosition(oLowerSeamPosition, oMarkVec, TEMPLATE_SEAM_MARK_LENGTH, True)
    
    Dim oLowerSeam  As Object
    Set oLowerSeam = oMfgTemplateReport.GetSeamEntity(TemplateLowerSeam)
    
    Dim oGeomMoniker    As IMoniker
    If Not oLowerSeam Is Nothing Then
        Set oGeomMoniker = oMfgRuleHelper.GetMoniker(oLowerSeam)
    End If
    
    Dim oMfgGeom3D As IJMfgGeom3d
    Set oMfgGeom3D = CreateGeom3D(oMarkCS, STRMFG_SEAM_MARK, oGeomMoniker, "Seam Mark")
    
    oGeomElems.Add oMfgGeom3D
    
    Set oMarkCS = Nothing
    Set oMarkVec = Nothing
    Set oGeomMoniker = Nothing
    Set oMfgGeom3D = Nothing
    
    Dim oUpperSeamPosition    As IJDPosition
    Set oUpperSeamPosition = oMfgTemplateReport.GetPoint(UpperSeamPoint)
    
    Set oMarkVec = GetMarkVector(oMfgTemplate, oTemplateBottomCurve, oUpperSeamPosition)
    Set oMarkCS = CreateMarkAtPosition(oUpperSeamPosition, oMarkVec, TEMPLATE_SEAM_MARK_LENGTH, True)
        
    Dim oUpperSeam  As Object
    Set oUpperSeam = oMfgTemplateReport.GetSeamEntity(TemplateUpperSeam)
    
    If Not oUpperSeam Is Nothing Then
        Set oGeomMoniker = oMfgRuleHelper.GetMoniker(oUpperSeam)
    End If
    
    Set oMfgGeom3D = CreateGeom3D(oMarkCS, STRMFG_SEAM_MARK, oGeomMoniker, "Seam Mark")
    
    oGeomElems.Add oMfgGeom3D
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oGeomElems
    
    Exit Function

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3021, , "RULES")
End Function
