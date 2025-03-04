VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompartFactors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: CompartFactors
'
'  History:
'  AppaRao            31st Mar 2005             Creation
'***************************************************************************
Option Explicit

Private Const MODULE = "CompartFactors: "
Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259
Implements IJCustomCompute

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

Private Sub IJCustomCompute_Evaluate(ByVal pCompartEntity As Object)
    On Error GoTo ErrorHandler
    
    SetValueToAttribute pCompartEntity
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, "CustomCompute::IJCustomCompute_Evaluate", Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub SetValueToAttribute(oObj As Object)
    On Error GoTo ErrorHandler

    Dim oAttributeMetadata          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttributesCollection       As IJDAttributesCol
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    ' Cache the interface iids
    Set oAttributeMetadata = oObj
    Set oAttrHelper = oObj
    
    Set oInterfaceInfo = oAttributeMetadata.InterfaceInfo(oAttributeMetadata.iid("IJUACompartFactors"))
    
    Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
    
    On Error Resume Next
    For Each oAttribute In oAttributesCollection
        Select Case oAttribute.AttributeInfo.Name
        Case "SteelReduction"
            oAttribute.value = GetSteelReduction()
        Case "Permeability"
            oAttribute.value = GetPermeability()
        Case "Tmy"
            oAttribute.value = GetTmy()
        Case "Capacity"
            oAttribute.value = GetCapacity()
        End Select
    Next
    On Error GoTo ErrorHandler
    
    Set oAttributeMetadata = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, "CustomCompute::SetValueToAttribute", Err.Description
    Err.Raise E_FAIL
End Sub
Private Function GetSteelReduction() As Double
Const METHOD = "GetSteelReduction"
On Error GoTo ErrorHandler

    GetSteelReduction = 100
    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompartGeneric::GetSteelReduction", Err.Description
    Err.Raise E_FAIL
End Function
Private Function GetPermeability() As Double
Const METHOD = "GetPermeability"
On Error GoTo ErrorHandler

    GetPermeability = 100
    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompartGeneric::GetPermeability", Err.Description
    Err.Raise E_FAIL
End Function
Private Function GetTmy() As Double
Const METHOD = "GetTmy"
On Error GoTo ErrorHandler

    GetTmy = 100
    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompartGeneric::GetTmy", Err.Description
    Err.Raise E_FAIL
End Function

Private Function GetCapacity() As Double
Const METHOD = "GetCapacity"
On Error GoTo ErrorHandler

    GetCapacity = 100
    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompartGeneric::GetCapacity", Err.Description
    Err.Raise E_FAIL
End Function
