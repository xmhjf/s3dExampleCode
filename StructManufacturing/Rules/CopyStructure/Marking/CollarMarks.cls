VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2009, Intergraph Corp.  All rights reserved.
'
' Project: MfgCSOptions
' Module: CollarMarks
'
' Description:  Create CollarMarks on the copied Plate
'
' Author: Manasa.J
'
' Comments:
' User         Date        Comments
' Manasa.J     9.12.09     Initial creation
'*******************************************************************************
Option Explicit
Private Const MODULE = "MfgCSOptions.CollarMarks"

Implements IJDMfgCopyStructureMarksRule
Private Function IJDMfgCopyStructureMarksRule_CreateMarks(ByVal oInputPlates As IJElements, ByVal oNewPlates As IJElements) As IJElements

    Const METHOD = "IJDMfgCopyStructureMarksRule_CreateMarks"
    On Error GoTo ErrorHandler
    
    Set IJDMfgCopyStructureMarksRule_CreateMarks = CreateMarksOfGivenType(oInputPlates, oNewPlates, STRMFG_COLLARPLATELOCATION_MARK)
            
    Exit Function

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1007, , "RULES")
End Function

