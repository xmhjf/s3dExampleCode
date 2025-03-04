VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadCompute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: CustomLoad
'
'  Abstract: The file contains an implementation of the default LoadCompute
'            for the CustomLoadObject.
'
'  History:
'  Raman            21th April 2004          Creation
'***************************************************************************
Option Explicit

Private Const MODULE = "LoadCompute: "

Implements IJCustomLoadCompute

Private Sub IJCustomLoadCompute_Evaluate(ByVal oLoad As GSCADSpatialLoadSemantics.IJLoad)
Const METHOD = "IJCustomLoadCompute_Evaluate"
On Error GoTo Errorhandler

    Dim oSpatialLoadOperation   As IJSpatialLoadOperation
    Dim oSpatialLoadGenericUtil As IJSpatialLoadGenericUtil
    Dim oCompartment            As IJCompartEntity
    Dim oCompartAttributes      As IJCompartAttributes
    Dim oLoadWCG                As IJWeightCG
    Dim dDensity                As Double
    Dim dVolume                 As Double
    Dim dWeight                 As Double
        
    'create load utility
    Set oSpatialLoadGenericUtil = New SpatialLoadGenericUtil
    
    'get the density of the load
    dDensity = oSpatialLoadGenericUtil.GetDesiredOccurrenceAttributeValue(oLoad, "density")
    
    'from load get the geometry and calculate volume
    Set oSpatialLoadOperation = oLoad
    
    Set oCompartment = oSpatialLoadOperation.Compartment
    Set oCompartAttributes = oCompartment
    dVolume = oCompartAttributes.Volume
    
    'calculate the load
    dWeight = dDensity * dVolume
    Set oLoadWCG = oLoad
    
    oLoadWCG.SetWCG WCG_WET, dWeight, oCompartAttributes.CogX, oCompartAttributes.CogY, oCompartAttributes.CogZ
    oLoadWCG.SetWCG WCG_DRY, dWeight, oCompartAttributes.CogX, oCompartAttributes.CogY, oCompartAttributes.CogZ

Exit Sub
Errorhandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_LOADCOMPUTE_FAILED_LOAD)
End Sub
