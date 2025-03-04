VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2011 Intergraph Corporation.  All Rights Reserved.
'
' File: CustomInputs.cls
'
' Author: Siva
'
' Abstract: Mfg Service Manager Custom inputs rule
'
'  History:
'   11th JULY 2011  Siva     Creation
'********************************************************************

Private Const MODULE = "MfgDepObjs"
Implements IJDMfgServiceMgrCustomInputs

Private Sub IJDMfgServiceMgrCustomInputs_GetMfgCustomInputSettings(lMarkingLine As Long, lMargin As Long, lPartShrinkage As Long, lAssemblyPartShrinkage As Long, lAssemblyShrinkage As Long, lTemplate As Long, lPinJig As Long, lMfgPlate As Long, lMfgProfile As Long)
    Const METHOD = "IJDMfgServiceMgrCustomInputs_GetMfgCustomInputSettings"
    On Error GoTo ErrorHandler
    
    ' So far, the service manager has been checking it's up-to-date status related to modeling inputs, like part and it's
    ' connection information and manufacturing definition objects (marking lines, margin, shrinkage);
    ' Based on the rule requirements within several Mfg objects,
    ' we need to have track additional objects, special related to planning related entities, which can be an input to the XML.
    ' Therefore is there a need for customizing the MSM to satisfy the requirements of different customers.
    ' Input settings can sent as combination also as shown below for Mfg Plate and Mfg Profile.
        
        ' 1.  StructMfgServiceMgr_NoCustomInputs  --> No custom inputs
        ' 2.  StructMfgServiceMgr_CustomRule  --> Rule will provide custom input collection
        ' 3.  StructMfgServiceMgr_PartProdRouting --> Production Routing
        ' 4.  StructMfgServiceMgr_PartParentAssembly --> Assembly
        ' 5.  StructMfgServiceMgr_PartCommonPart  --> Common Parts Group
        ' 6.  StructMfgServiceMgr_PartKnuckles  --> Profile Knuckle (Split, vs bend)
        ' 7.  StructMfgServiceMgr_PartFeatures  --> Features
        ' 8.  StructMfgServiceMgr_PartOpenings  --> Holes & openings
        ' 9.  StructMfgServiceMgr_AllAssemblyShrinkages  --> Assembly shrinkages on parent assemblies
        ' 10.  StructMfgServiceMgr_AssemblyBasePlate  --> Base plate of the assembly
        ' 11.  StructMfgServiceMgr_BasePlateStiffeners  --> Stiffeners on the plate or on the base plate of an assembly
        ' 12.  StructMfgServiceMgr_BUAssemblyPCs  --> Physical Connections in the built up assembly
    
    lMarkingLine = StructMfgServiceMgr_NoCustomObjects  ' No custom inputs
    lMargin = StructMfgServiceMgr_NoCustomObjects       ' No custom inputs
    lTemplate = StructMfgServiceMgr_NoCustomObjects     ' No custom inputs
    lPinJig = StructMfgServiceMgr_NoCustomObjects       ' No custom inputs
    
    lPartShrinkage = StructMfgServiceMgr_NoCustomObjects    ' No custom inputs
    lAssemblyPartShrinkage = StructMfgServiceMgr_NoCustomObjects    ' No custom inputs
    lAssemblyShrinkage = StructMfgServiceMgr_NoCustomObjects    ' No custom inputs
    
    ' Mfg Plate - Openings as custom inputs
    lMfgPlate = StructMfgServiceMgr_PartOpenings + StructMfgServiceMgr_PartFeatures
    
    '  Mfg profile - Openings as custom inputs
    lMfgProfile = StructMfgServiceMgr_PartOpenings + StructMfgServiceMgr_PartFeatures
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function IJDMfgServiceMgrCustomInputs_GetMfgCustomDependencyCollection(ByVal oMfgObject As Object) As IMSCoreCollections.IJElements
    Const METHOD = "IJDMfgServiceMgrCustomInputs_GetMfgCustomInputsCollection"
    On Error GoTo ErrorHandler
    
    ' Code to get the rule based custom inputs for the given Mfg object
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
