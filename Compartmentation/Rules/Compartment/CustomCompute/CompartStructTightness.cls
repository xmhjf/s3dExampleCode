VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompartStructTightness"
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
'  Raghuveer            June 7th 2006             Creation
'***************************************************************************
Option Explicit

Private Const MODULE = "StructTightNess: "
Private Const STRUCTTIGHTNESSATTRIBUTE = "StructTightness"
Private Const COMPARTTIGHTNESSATTRIBUTE = "CompartTightness"
Private Const IJCOMPARTTIGHTNESS = "IJCompartTightness"
Implements IJCustomCompute

Private Sub IJCustomCompute_Evaluate(ByVal pCompartEntity As Object)
On Error GoTo ErrorHandler
Const METHOD = "IJCustomCompute_Evaluate"
    
    SetTightnessValue pCompartEntity
    
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

Private Sub SetTightnessValue(oObj As Object)
On Error GoTo ErrorHandler
Const METHOD = "SetTightnessValue"
    
    Dim oAttributeMetadata          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttributesCollection       As IJDAttributesCol
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    Dim lStructTightness            As StructPlateTightness

    Set oAttributeMetadata = oObj
    Set oAttrHelper = oObj
           
    Set oInterfaceInfo = oAttributeMetadata.InterfaceInfo(oAttributeMetadata.iid(IJCOMPARTTIGHTNESS))
    
    If Not oInterfaceInfo Is Nothing Then

        Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
        
        If Not oAttributesCollection Is Nothing Then
        
            For Each oAttribute In oAttributesCollection
                If oAttribute.AttributeInfo.Name = STRUCTTIGHTNESSATTRIBUTE Then
                
                    lStructTightness = GetMinTightness(oObj)
                    If ((lStructTightness >= UnSpecifiedTightness) And (lStructTightness <= AirTight)) Then
                        oAttribute.Value = lStructTightness
                    End If
                End If
            Next
                        
            For Each oAttribute In oAttributesCollection
                If oAttribute.AttributeInfo.Name = COMPARTTIGHTNESSATTRIBUTE Then
                    If lStructTightness = UnSpecifiedTightness Then
                        oAttribute.Value = 1  'UnSpecified
                        Exit For
                    End If
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
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_STRUCTTIGHTNESS_FAILED_TIGHTNESSEVALUATION)
End Sub
