VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FixPointLightDuty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright   (c) 2008, Intergraph Corporation.  All Rights reserved.
'
'   FixPointLightDuty.cls
'   ProgID:         HS_Hilti_MQAssy.FixPointLightDuty
'   Author:         JOGN
'   Createion Date: October 8, 2008
'   Description:    Fixed Point Set assembly for Hilti MQ library
'
'   Change History:
'   Date            who         change description
'   16-02-2010      VSP         Addition to S3D
'   09-11-2011      AS          TR-CP-193768  Inconsistent placement observed with the Hilti Supports
'   04-04-2013      Ramya       CR 226479  FabRespons codelist should be removed from MQ Parts Codelist
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "FixPointLightDuty"

Implements IJHgrAssmInfo
Implements IJHgrSymbolBOMServices

Private numParts As Integer
Private nPlate As Integer
Private nRod As Integer
Private nPipeRing As Integer

Private sPlate As String
Private sRod As String
Private sPipeRing As String
Private dH As Double
Private lFabReq As Long
Private lFabType As Long
Private lConReq As Long
Private lConType As Long
Private lFabRespons As Long
Private sPartNum As String



'----------------------------------------------------------------------------------------------
'   Name:           IJHgrAssmInfo_GetAssemblyCatalogParts()
'   Description:    Will Define part class and select the parts.
'   Date - Author:  October 3, 2008 - JOGN
'
'   Inputs:         pDispInputConfigHlpr as Object
'   Outputs:        A collection of object implementing IJDPart interface
'----------------------------------------------------------------------------------------------

Private Function IJHgrAssmInfo_GetAssemblyCatalogParts(ByVal pDispInputConfigHlpr As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetAssemblyCatalogParts"

    On Error GoTo ErrorHandler
    
    Initialize pDispInputConfigHlpr

    nPlate = 1
    nRod = 2
    nPipeRing = 3

    sPlate = GetAttr("Plate")
    sRod = GetAttr("Rod")
    sPipeRing = GetAttr("PipeRing")
    dH = GetAttr("H")
    lFabReq = GetAttr("FabricationRequirement")
    lFabType = GetAttr("FabricationType")
    lConReq = GetAttr("ConstructionRequirement")
    lConType = GetAttr("ConstructionType")
    lFabRespons = GetAttr("FabricationResponsibility")

    sPartNum = GetAttr("PartNum")

    Hilti_RouteSizeIsValid sPartNum

    numParts = 3
    ReDim PartClasses(numParts) As String
    PartClasses(nPlate) = sPlate
    PartClasses(nRod) = sRod
    PartClasses(nPipeRing) = sPipeRing
    
    Set IJHgrAssmInfo_GetAssemblyCatalogParts = AddPartsByName(PartClasses)

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Function IJHgrAssmInfo_GetAssemblyJoints(ByVal pDispInputConfigHlpr As Object, _
                                                 ByVal pDispPartOccCollection As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetAssemblyJoints"

    On Error GoTo ErrorHandler:

    Dim myIJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set myIJHgrInputConfigHlpr = pDispInputConfigHlpr

    'Get interface for accessing items on the collection of Part Occurences
    Dim IJElements_PartOccCollection As IJElements
    Set IJElements_PartOccCollection = pDispPartOccCollection

    InitializeJoints pDispPartOccCollection

    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationRequirement", IJElements_PartOccCollection.Item(nPipeRing), lFabReq
    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationType", IJElements_PartOccCollection.Item(nPipeRing), lFabType
    myIJHgrInputConfigHlpr.SetAttributeValue "ConstructionRequirement", IJElements_PartOccCollection.Item(nPipeRing), lConReq
    myIJHgrInputConfigHlpr.SetAttributeValue "ConstructionType", IJElements_PartOccCollection.Item(nPipeRing), lConType
    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationResponsibility", IJElements_PartOccCollection.Item(nPipeRing), lFabRespons

    ' Create a collection to hold the joints
    Dim JointCollection As New Collection
    Dim JointFactory As New HgrJointFactory
    Dim AssemblyJoint As Object

    Dim sTopPlatePort As String
    Dim sBottomPlatePort As String
    Dim iPlateRodConfigIdx As Integer, iPlateStructConfigIdx As Integer

    If sPlate = "Hilti MFP-GP 20-F" Then
        sTopPlatePort = "Top"
        sBottomPlatePort = "Bottom"
        iPlateRodConfigIdx = 9380
        iPlateStructConfigIdx = 9380
    ElseIf sPlate = "Hilti MRG 2-F" Then
        sTopPlatePort = "Base"
        sBottomPlatePort = "Rod"
        iPlateRodConfigIdx = 9444
        iPlateStructConfigIdx = 9444
    End If

    If Hilti_StructIsSloped() Then
        PF_EventHandler "Assembly can not be placed to sloped Slab/Structure.", Err, MODULE, METHOD, False
    End If

    'Add Joints
    AddJoint JointFactory.MakePlanarJoint(nPlate, sTopPlatePort, -1, "Structure", iPlateStructConfigIdx)
    AddJoint JointFactory.MakeRigidJoint(nPlate, sBottomPlatePort, nRod, "RodEnd1", iPlateRodConfigIdx)
    AddJoint JointFactory.MakePrismaticJoint(nRod, "RodEnd1", nRod, "RodEnd2", 3574)
    AddJoint JointFactory.MakeRevoluteJoint(nRod, "RodEnd2", nPipeRing, "Side", 82)
    AddJoint JointFactory.MakeRigidJoint(nPipeRing, "Route", -1, "Route")

    ' Return the collection of Joints
    Set IJHgrAssmInfo_GetAssemblyJoints = GetJoints

    Class_Terminate

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Function IJHgrAssmInfo_GetMaxRouteConnectionValue() As Long

    IJHgrAssmInfo_GetMaxRouteConnectionValue = 1

End Function

Private Function IJHgrAssmInfo_GetRouteConnections(ByVal pDispInputConfigHlpr As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetRouteConnections"

    On Error GoTo ErrorHandler

    Set IJHgrAssmInfo_GetRouteConnections = ConnectPartToRouteOrStruct(nPipeRing, 1)   'partindex, routeindex

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Function IJHgrAssmInfo_GetStructConnections(ByVal pDispInputConfigHlpr As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetStructConnections"

    On Error GoTo ErrorHandler

    Set IJHgrAssmInfo_GetStructConnections = ConnectPartToRouteOrStruct(nPlate, 1)

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Sub IJHgrSymbolBOMServices_EvaluateBOM(ByVal pSupportComp As Object, bstrBOMDesc As String)
    Const METHOD = "IJHgrSymbolBOMServices_EvaluateBOM"

    On Error GoTo ErrorHandler

    Initialize pSupportComp

    bstrBOMDesc = "Hilti"

    Class_Terminate

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub




