VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegionCorrosionRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: CompartStructTightness
'  ProgId:  CompartSemantics.CompartStructTightness
'
'  History:
'  Raghuveer            July 21st 2006             Creation
'***************************************************************************
Option Explicit

Private Const MODULE = "RegionCorrosionRule "
Private Const IJUAREGIONCORROSIONRULE = "IJUARegionCorrosionRule"
Private Const CORROSIONRULEATTRIBUTE = "CorrosionRule"
Implements IJCustomCompute

Private Sub IJCustomCompute_Evaluate(ByVal pCompartEntity As Object)
On Error GoTo ErrorHandler
Const METHOD = "IJCustomCompute_Evaluate"

        SetCorrosionValue pCompartEntity
    
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

Private Sub SetCorrosionValue(oObj As Object)
On Error GoTo ErrorHandler
Const METHOD = "SetCorrosionValue"
    
    Dim oAttributeMetadata          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttributesCollection       As IJDAttributesCol
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo

    Set oAttributeMetadata = oObj
    Set oAttrHelper = oObj
                    
    Set oInterfaceInfo = oAttributeMetadata.InterfaceInfo(oAttributeMetadata.iid(IJUAREGIONCORROSIONRULE))
        
    
    If Not oInterfaceInfo Is Nothing Then

        Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
            
        
        If Not oAttributesCollection Is Nothing Then
        
            For Each oAttribute In oAttributesCollection
                If oAttribute.AttributeInfo.Name = CORROSIONRULEATTRIBUTE Then
                    oAttribute.Value = 1234
                End If
            Next
            
        End If
    End If
    
    
    Set oAttributeMetadata = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing
        
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub
