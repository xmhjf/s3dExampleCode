VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TmpMrkSltCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project: MfgTemplateMarking
'
'  Abstract:    Creates Template marks on Template
'
'  History:
'      Siva        2nd September 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingPlate.TmpMrkSltCut"

Implements IJMfgTemplateMarkingRule
Implements IJMfgTemplateSlotParameters

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler

    Dim oMfgTemplateReport  As IJMfgTemplateReport
    Set oMfgTemplateReport = oMfgTemplate
    
    ' Get the template marks from the Template report
    Dim oGeomElems As IJElements
    Set oGeomElems = oMfgTemplateReport.GetMarks(STRMFG_TEMPLATE_MARK)

    Set IJMfgTemplateMarkingRule_CreateMarks = oGeomElems
    Exit Function

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3024, , "RULES")
End Function

Private Sub IJMfgTemplateSlotParameters_GetSlotParameters(lSlotGeometryType As Long, lSlotAtTop As Long, lSlotAtEdgeType As Long, lSlotLocationType As Long, dSlotMarginAtLower As Double, dSlotMarginAtUpper As Double, dSlotMarginAtBottom As Double)
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler

    lSlotGeometryType = MfgTemplate_SlotAsCut 'SlotAsCut
    lSlotAtTop = MfgTemplate_TopSlotForPrimary
    lSlotAtEdgeType = MfgTemplate_EdgeSlotAsSlot
    lSlotLocationType = MfgTemplate_SlotAtCenter
    dSlotMarginAtLower = 0 '0mm
    dSlotMarginAtUpper = 0 '0mm
    dSlotMarginAtBottom = 0 '0mm
    
    Exit Sub

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3024, , "RULES")
End Sub
