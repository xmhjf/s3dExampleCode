VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFeatureType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CFeatureType.cls
'   ProgID:         HgrFeatureType.CFeatureType
'   Author:         Sundar
'   Creation Date:  20.Feb.2002
'   Description:
'
'
'   Change History:
'       20.Feb.2002             Sundar       Creation
'       12.Apr.2002             Caroline     Fix for TR 24565
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Implements IJHgrSupportRule

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259

Private Const MODULE = "HgrAttributesRules::CFeatureType:: "
Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub
Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

Private Function IJHgrSupportRule_AttributeValue(ByVal pdispInputConfigHlpr As Object) As Variant()
    Const METHOD = "IJHgrSupportRule_AttributeValue"
    On Error GoTo ErrorHandler
    
    Dim oInputConfigHlpr As IJHgrInputConfigHlpr
    Dim oInputObjectInfo As IJHgrInputObjectInfo
    Dim oRtePathFeat As IJRtePathFeat
    Dim ePathFeatureObjectTypes As PathFeatureObjectTypes
    
    Dim varFeatureType(0 To 0) As Variant
    varFeatureType(0) = "UnKnown"
    
    On Error Resume Next
    ' check for input object info
    Set oInputObjectInfo = pdispInputConfigHlpr
    On Error GoTo ErrorHandler
    
    If Not oInputObjectInfo Is Nothing Then
        ' get the pipe collection
        Dim oPipeCollection As IJElements
        Set oPipeCollection = oInputObjectInfo.GetSupportedObjects()
        
        ' find the feature type
        On Error Resume Next
        Set oRtePathFeat = oPipeCollection(1)
        On Error GoTo ErrorHandler
        
        ' get the feature type
        ePathFeatureObjectTypes = oRtePathFeat.GetPathFeatureObjectType
        Select Case ePathFeatureObjectTypes
            Case PathFeatureType_STRAIGHT, PathFeatureType_CURVE
                varFeatureType(0) = "STRAIGHT"
            Case PathFeatureType_TURN
                varFeatureType(0) = "TURN"
            Case PathFeatureType_ALONG
                varFeatureType(0) = "PART"
            Case PathFeatureType_END
                varFeatureType(0) = "END"
            Case PathFeatureType_SURFACE
                varFeatureType(0) = "SURFACE"
            Case Else
                varFeatureType(0) = "UnKnown"
        End Select
        Set oPipeCollection = Nothing
        Set oRtePathFeat = Nothing
    End If
        
    ' release the objects
    Set oInputObjectInfo = Nothing
    
    ' return the feature types
    IJHgrSupportRule_AttributeValue = varFeatureType
    
    Exit Function
ErrorHandler:
Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
