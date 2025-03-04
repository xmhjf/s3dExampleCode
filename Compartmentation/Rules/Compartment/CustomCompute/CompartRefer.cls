VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompartRefer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: CustomLoad
'
'  Abstract: The file contains an implementation of the Custom Compute - which gets processed during modification/creation of volume.
'           This code is presented as an example to do the following:
'       --- Update IJNamedItem so that the name rule (if associated) gets triggerred
'       --- Set the attribute "ReferencePosition" of the virtual interface "IJCompartReference"
'
'  History:
'  Arnab            16th Sept 2004              Creation
'  AppaRao          19th April 2005             LogError method has been used to log th error
'***************************************************************************
Option Explicit

Private Const MODULE = "CustomCompute: "
Private Const IID_IJNAMEDITEM = "{7330A824-7897-4fc5-A279-A01E5476ECDD}"

Implements IJCustomCompute

'This is the method called at final stages of Compute.
'This is implemented to do the following:
'       --- Update IJNamedItem so that the name rule (if associated) gets triggerred
'       --- Set the attribute "ReferencePosition" of the virtual interface "IJCompartReference"
Private Sub IJCustomCompute_Evaluate(ByVal pCompartEntity As Object)
On Error GoTo ErrorHandler
Const METHOD = "IJCustomCompute_Evaluate"

    Dim oCompartNamedItem       As IJNamedItem
    Dim oCompartEntity          As IJCompartEntity
    
    'Dirty the interface IJNamedItem
    Set oCompartEntity = pCompartEntity
    oCompartEntity.DirtyInterface IID_IJNAMEDITEM
    Set oCompartEntity = Nothing
    
    'Set the attribute
    On Error Resume Next
    SetValueToAttribute pCompartEntity
    
    Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
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
    
    Set oInterfaceInfo = oAttributeMetadata.InterfaceInfo(oAttributeMetadata.iid("IJCompartReference"))
    
    Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
    
    For Each oAttribute In oAttributesCollection
        If oAttribute.AttributeInfo.Name = "ReferencePosition" Then
            oAttribute.Value = GetReferencePosition(oObj)
        End If
    Next
   
    Set oAttributeMetadata = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_REFERENCE_FAILED_REFERENCEPOSITION)
End Sub
