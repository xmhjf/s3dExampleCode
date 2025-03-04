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
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Frames And Edges Position Process rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit
Private Const MODULE = "MfgTemplateProcess.PositionFrame"
Private Const PART_SURFACE = 0
Private Const MOLDED_SURFACE = 1
Private Const PART_SURFACE_BASE = 2

Implements IJMfgTemplatePositionRule
 
Private Sub IJMfgTemplatePositionRule_GetPositionData(ByVal pFrameSysDsp As Object, ByVal pPartDisp As Object, ByVal pProcessSettings As Object, ByVal pMfgTemplateSet As Object, dEdgeOffset As Double, lNumOfTemplates As Long, bSupportEdges As Boolean, ppFrames As IMSCoreCollections.IJElements)
Const METHOD = "IJMfgTemplatePositionRule_GetPositionData"
On Error GoTo ErrorHandler

    
    Set ppFrames = GetPositionFrames(pFrameSysDsp, pPartDisp, pProcessSettings, pMfgTemplateSet)
    
    dEdgeOffset = 0
    lNumOfTemplates = 0
    bSupportEdges = True

 Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3010, , "RULES")
End Sub

