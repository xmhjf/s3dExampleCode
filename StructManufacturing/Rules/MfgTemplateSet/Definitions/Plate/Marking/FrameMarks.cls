VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project: MfgTemplateMarking
'
'  Abstract:    Create Frame control line(Frame/Buttock/Water) marks on Template
'
'  History:
'      Siva        2nd September 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingPlate.FrameMark"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler

    Dim oGeomElems As IJElements
    Set oGeomElems = New JObjectCollection
    
    Dim oMfgTemplateReport  As IJMfgTemplateReport
    Set oMfgTemplateReport = oMfgTemplate
    
    Dim oTempColl  As IJElements
    
    ' Get the frame line marks from the Template report
    Set oTempColl = oMfgTemplateReport.GetMarks(STRMFG_FRAMELINE_MARK)
    oGeomElems.AddElements oTempColl
    Set oTempColl = Nothing
    
    ' Get the buttock line marks from the Template report
    Set oTempColl = oMfgTemplateReport.GetMarks(STRMFG_BUTTOCKLINE_MARK)
    oGeomElems.AddElements oTempColl
    Set oTempColl = Nothing
    
    ' Get the water line marks from the Template report
    Set oTempColl = oMfgTemplateReport.GetMarks(STRMFG_WATERLINE_MARK)
    oGeomElems.AddElements oTempColl
    Set oTempColl = Nothing
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oGeomElems
    Exit Function
    
ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3018, , "RULES")
End Function
