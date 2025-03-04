VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AverageShrinkages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2013, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgShrinkageRules
' Module: MultiplyShrinkages
'
' Description:  This rule accumulates multiple shrinkages by averaging them.
'
' Author:
'
' Comments:
' July 29th, 2015    StructMfg     New desgin of the class
'*******************************************************************************
Option Explicit

Const MODULE = "StrMfgShrinkageRules:MultiplyShrinkages"

Implements IJDMfgShrinkageAccumulationRule
Private Sub IJDMfgShrinkageAccumulationRule_AccumulateMultipleShrinkages(ByVal pPartDisp As Object, ByVal pMfgShrinkageColl As Object, ByVal pPrimaryDir As AutoMath.IJDVector, ByVal pSecondaryDir As AutoMath.IJDVector, pdPrimaryFactor As Double, pdSecondaryFactor As Double)
    Const METHOD = "IJDMfgShrinkageAccumulationRule_AccumulateMultipleShrinkages"
    On Error GoTo ErrorHandler
    
    AccumulateShrinkageParameters pMfgShrinkageColl, pPrimaryDir, pSecondaryDir, _
        AVERAGE_SHRINKAGES, pdPrimaryFactor, pdSecondaryFactor
        
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub




