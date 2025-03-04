VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PosFrameEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgTemplateProcess
' Module:
'
' Description:  Determines the process settings for the Template
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
' 2006.09.11    Ram     Modified to use MfgFrameSystem
'*******************************************************************************
Option Explicit
Private Const MODULE = "MfgTemplateProcess.PositionFrame"
Private Const PART_SURFACE = 0
Private Const MOLDED_SURFACE = 1
Private Const PART_SURFACE_BASE = 2

Implements IJMfgTemplatePositionRule
 
Private Sub IJMfgTemplatePositionRule_GetPositionData(ByVal pFrameSysDsp As Object, ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart, ByVal pProcessSettings As Object, ByVal pMfgTemplateSet As Object, dEdgeOffset As Double, lNumOfTemplates As Long, bSupportEdges As Boolean, ppFrames As IMSCoreCollections.IJElements)
Const METHOD = "IJMfgTemplatePositionRule_GetPositionData"
On Error GoTo ErrorHandler

    
    Set ppFrames = GetPositionFrames(pFrameSysDsp, pPlatePart, pProcessSettings, pMfgTemplateSet)
    
    dEdgeOffset = 0
    lNumOfTemplates = 0
    bSupportEdges = True

 Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3010, , "RULES")
End Sub

