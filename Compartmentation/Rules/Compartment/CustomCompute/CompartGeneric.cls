VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompartGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: CompartGeneric
'
'  History:
'  AppaRao            31st Mar 2005             Creation
'***************************************************************************
Option Explicit

Private Const MODULE = "CompartGeneric: "
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
    m_oErrors.Add Err.Number, "CompartGeneric::IJCustomCompute_Evaluate", Err.Description
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
    
    Set oInterfaceInfo = oAttributeMetadata.InterfaceInfo(oAttributeMetadata.iid("IJUACompartGeneric"))
    
    Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
    
    On Error Resume Next
    For Each oAttribute In oAttributesCollection
        Select Case oAttribute.AttributeInfo.Name
            Case "Description"
                oAttribute.value = GetDescription()
            Case "Purpose"
                oAttribute.value = GetPurpose()
            Case "PurposeDescription"
                oAttribute.value = GetPurposeDescription()
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
    m_oErrors.Add Err.Number, "CompartGeneric::SetValueToAttribute", Err.Description
    Err.Raise E_FAIL
End Sub

Private Function GetDescription() As String
Const METHOD = "GetDescription"
On Error GoTo ErrorHandler

    GetDescription = "Description"
    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompartGeneric::GetDescription", Err.Description
    Err.Raise E_FAIL
End Function
Private Function GetPurpose() As String
Const METHOD = "GetPurpose"
On Error GoTo ErrorHandler

    GetPurpose = "Purpose"
    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompartGeneric::GetPurpose", Err.Description
    Err.Raise E_FAIL
End Function
Private Function GetPurposeDescription() As String
Const METHOD = "GetPurposeDescription"
On Error GoTo ErrorHandler

    GetPurposeDescription = "Purposedescription"
    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompartGeneric::GetPurposeDescription", Err.Description
    Err.Raise E_FAIL
End Function



