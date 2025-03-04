VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SightLnMrk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgTemplateMarking
' Module: SightLine
'
' Description:  Determines the sighline for a template
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit
Private Const MODULE = "MfgTemplateMarking.SightLine"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler
    
    Dim oTemplateTopLine   As IJLine
    Set oTemplateTopLine = oMfgTemplate.TopLine
    
    If oTemplateTopLine Is Nothing Then
        Exit Function
    End If
    
    Dim oMfgGeomChild   As IJMfgGeomChild
    Set oMfgGeomChild = oMfgTemplate
    
    Dim oMfgTemplateSet As IJDMfgTemplateSet
    Set oMfgTemplateSet = oMfgGeomChild.GetParent
    
    Dim oTemplateTopLineCS   As IJComplexString
    Set oTemplateTopLineCS = New ComplexString3d
    oTemplateTopLineCS.AddCurve oTemplateTopLine, False
    
    Dim oMfgTemplateReport  As IJMfgTemplateReport
    Set oMfgTemplateReport = oMfgTemplate
    
    Dim oBCTLPosition    As IJDPosition
    Set oBCTLPosition = oMfgTemplateReport.GetPoint(BaseControlTopLinePoint)
    
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New Helper
    
    Dim oTopLineWB  As IJWireBody
    Set oTopLineWB = oMfgRuleHelper.ComplexStringToWireBody(oTemplateTopLineCS)
    
    Dim oStartPos   As IJDPosition, oEndPos     As IJDPosition
    oTopLineWB.GetEndPoints oStartPos, oEndPos
    
    Dim oTopLineVec As IJDVector
    Set oTopLineVec = oEndPos.Subtract(oStartPos)
     
    Dim oUpDirVec As IJDVector
    oMfgTemplateSet.GetDirectionVector Nothing, oUpDirVec, oMfgTemplate.GroupIndex
    
    ' Mark position is offset from the BCTL positin
    Dim oMarkPosition    As IJDPosition
    
    ' Make sure sight line mark is on the right side of the BCTL
    If oUpDirVec.Dot(oTopLineVec) > 0 Then
        Set oMarkPosition = oMfgRuleHelper.GetPointAlongCurveAtDistance(oTopLineWB, oBCTLPosition, TEMPLATE_SIGHT_LINE_OFFSET, oEndPos)
    Else
        Set oMarkPosition = oMfgRuleHelper.GetPointAlongCurveAtDistance(oTopLineWB, oBCTLPosition, TEMPLATE_SIGHT_LINE_OFFSET, oStartPos)
    End If
    
    Dim oMarkVector    As IJDVector
    Set oMarkVector = GetMarkVector(oMfgTemplate, oTemplateTopLine, oMarkPosition)
    
    Dim oMarkCS     As IJComplexString
    Set oMarkCS = CreateMarkAtPosition(oMarkPosition, oMarkVector, TEMPLATE_BCL_MARK_LENGTH, True)
    
    Dim oGeomElems As IJElements
    Set oGeomElems = New JObjectCollection
    
    
    Dim oBCL    As Object
    Set oBCL = oMfgTemplateSet.GetControlLine(oMfgTemplate.GroupIndex)
    
    Dim oGeomMoniker    As IMoniker
    Set oGeomMoniker = oMfgRuleHelper.GetMoniker(oBCL)
    
    Dim oMfgGeom3D As IJMfgGeom3d
    Set oMfgGeom3D = CreateGeom3D(oMarkCS, STRMFG_SIGHTLINE_MARK, oGeomMoniker, "Sight Line Mark")
    
    oGeomElems.Add oMfgGeom3D
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oGeomElems

    Exit Function

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 9999, , "RULES")
End Function
