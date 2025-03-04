VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnitLoadCompute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: CustomLoad
'
'  Abstract: The file contains an implementation of the default UnitLoadCompute
'            for the CustomLoadObject.
'
'  History:
'  Arnab                6th May 2005          Creation - copied from LoadCompute
'***************************************************************************
Option Explicit

Private Const MODULE = "UnitLoadCompute: "

Implements IJCustomLoadCompute

Private Sub IJCustomLoadCompute_Evaluate(ByVal oLoad As GSCADSpatialLoadSemantics.IJLoad)
Const METHOD = "IJCustomLoadCompute_Evaluate"
On Error GoTo Errorhandler

    Dim oUnitLoad               As IJUnitLoad
    Dim oLoadWCG                As IJWeightCG
    Dim xx                      As Double
    Dim yy                      As Double
    Dim zz                      As Double
    Dim dWeight                 As Double
        
    Set oUnitLoad = oLoad
    
    oUnitLoad.GetDirection xx, yy, zz
    
    'Weight is the -Z direction component of the load
    dWeight = oUnitLoad.Magnitude * (-zz)
    
    oUnitLoad.GetPosition xx, yy, zz
    Set oLoadWCG = oLoad
    
    oLoadWCG.SetWCG WCG_WET, dWeight, xx, yy, zz
    oLoadWCG.SetWCG WCG_DRY, dWeight, xx, yy, zz
    
Exit Sub
Errorhandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_UNITLOADCOMPUTE_FAILED_UNITLOAD)
End Sub
