VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompartFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: CompartFrame
'
'  History:
'  AppaRao            31st Mar 2005             Creation
'***************************************************************************
Option Explicit

Private Const MODULE = "CompartFrame "

Implements IJCustomCompute

Private Sub IJCustomCompute_Evaluate(ByVal pCompartEntity As Object)
On Error GoTo ErrorHandler
Const METHOD = "IJCustomCompute_Evaluate"
    
    SetValueToAttribute pCompartEntity
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

Private Sub SetValueToAttribute(oObj As Object)
On Error GoTo ErrorHandler
Const METHOD = "SetValueToAttribute"

    Dim oAttributeMetadata          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttributesCollection       As IJDAttributesCol
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    ' Cache the interface iids
    Set oAttributeMetadata = oObj
    Set oAttrHelper = oObj
    
    Set oInterfaceInfo = oAttributeMetadata.InterfaceInfo(oAttributeMetadata.iid("IJUACompartFrame"))
    
    Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
    
    If FramesExistsinModel = True Then
        For Each oAttribute In oAttributesCollection
            Select Case oAttribute.AttributeInfo.Name
                Case "LongitudinalMin"
                    oAttribute.Value = GetLongitudinalMin(oObj)
                Case "LongitudinalMax"
                    oAttribute.Value = GetLongitudinalMax(oObj)
                Case "DeckMin"
                    oAttribute.Value = GetDeckMin(oObj)
                Case "DeckMax"
                    oAttribute.Value = GetDeckMax(oObj)
                Case "TransversalMin"
                    oAttribute.Value = GetTransversalMin(oObj)
                Case "TransversalMax"
                    oAttribute.Value = GetTransversalMax(oObj)
            End Select
        Next
    End If

    Set oAttributeMetadata = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_FRAMES_FAILED_NAMES)
End Sub

