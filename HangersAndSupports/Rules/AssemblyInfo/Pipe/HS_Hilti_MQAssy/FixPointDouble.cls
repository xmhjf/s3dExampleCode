VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FixPointDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright   (c) 2008, Intergraph Corporation.  All Rights reserved.
'
'   FixPointDouble.cls
'   ProgID:         HS_Hilti_MQAssy.FixPointDouble
'   Author:         JOGN
'   Createion Date: October 8, 2008
'   Description:    Fixed Point Set assembly for Hilti MQ library
'
'   Change History:
'   Date            who         change description
'   05-02-2010      VSP         Addition to S3D
'   05-07-2010      VSP         Modified few lines of code for the plate part number
'   09-11-2011      AS          TR-CP-193768  Inconsistent placement observed with the Hilti Supports
'   04-04-2013      Ramya       CR 226479  FabRespons codelist should be removed from MQ Parts Codelist
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "FixPointDouble"

Implements IJHgrAssmInfo
Implements IJHgrSymbolBOMServices

Private numParts As Integer
Private nPlate As Integer
Private nRod As Integer
Private nRod2 As Integer
Private nPipeRing As Integer
Private nPipeRing2 As Integer

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

    nPlate = 1
    nRod = 2
    nRod2 = 3
    nPipeRing = 4
    nPipeRing2 = 5

    Initialize pDispInputConfigHlpr
    
    sPlate = GetAttr("Plate")
    sRod = GetAttr("Rod")
    sPipeRing = GetAttr("PipeRing")
    dH = GetAttr("H")
    sPartNum = GetAttr("PartNum")
    lFabReq = GetAttr("FabricationRequirement")
    lFabType = GetAttr("FabricationType")
    lConReq = GetAttr("ConstructionRequirement")
    lConType = GetAttr("ConstructionType")
    lFabRespons = GetAttr("FabricationResponsibility")

    Hilti_RouteSizeIsValid sPartNum

    numParts = 5
    ReDim PartClasses(numParts) As String
    PartClasses(nPlate) = sPlate
    PartClasses(nRod) = sRod
    PartClasses(nRod2) = sRod
    PartClasses(nPipeRing) = sPipeRing
    PartClasses(nPipeRing2) = sPipeRing

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

    ' Create a collection to hold the joints
    Dim JointCollection As New Collection
    Dim JointFactory As New HgrJointFactory
    Dim AssemblyJoint As Object

    Dim sTopPlatePort As String
    Dim sBottomPlatePort As String
    Dim sBottomPlatePort2 As String
    Dim dSlideL As Double
    Dim dRodOffset As Double
    Dim dSlideInset As Double

    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationRequirement", IJElements_PartOccCollection.Item(nPipeRing), lFabReq
    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationType", IJElements_PartOccCollection.Item(nPipeRing), lFabType
    myIJHgrInputConfigHlpr.SetAttributeValue "ConstructionRequirement", IJElements_PartOccCollection.Item(nPipeRing), lConReq
    myIJHgrInputConfigHlpr.SetAttributeValue "ConstructionType", IJElements_PartOccCollection.Item(nPipeRing), lConType
    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationResponsibility", IJElements_PartOccCollection.Item(nPipeRing), lFabRespons

    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationRequirement", IJElements_PartOccCollection.Item(nPipeRing2), lFabReq
    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationType", IJElements_PartOccCollection.Item(nPipeRing2), lFabType
    myIJHgrInputConfigHlpr.SetAttributeValue "ConstructionRequirement", IJElements_PartOccCollection.Item(nPipeRing2), lConReq
    myIJHgrInputConfigHlpr.SetAttributeValue "ConstructionType", IJElements_PartOccCollection.Item(nPipeRing2), lConType
    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationResponsibility", IJElements_PartOccCollection.Item(nPipeRing2), lFabRespons

    dSlideL = Hilti_MultipleInterfaceDataLookUp("SlideL", "JUAMQSlidePlate", "JCUHASMQ_DoubleRollConnHDG", "PartNumber", "JDPart", "'Hilti MRG-D6-F'")
    dRodOffset = Hilti_MultipleInterfaceDataLookUp("RodOffset", "JOAMQROffset", "JCUHASMQ_DoubleRollConnHDG", "PartNumber", "JDPart", "'Hilti MRG-D6-F'")
    dSlideInset = Hilti_MultipleInterfaceDataLookUp("SlideInset", "JUAMQSlidePlate", "JCUHASMQ_DoubleRollConnHDG", "PartNumber", "JDPart", "'Hilti MRG-D6-F'")

    sTopPlatePort = "Base"
    sBottomPlatePort = "Rod"
    sBottomPlatePort2 = "Rod2"

    If Hilti_StructIsSloped() Then
        PF_EventHandler "Assembly can not be placed to sloped Slab/Structure.", Err, MODULE, METHOD, False
    End If

    'Add Joints
    AddJoint JointFactory.MakeRigidJoint(nPlate, sBottomPlatePort, nRod, "RodEnd1")
    AddJoint JointFactory.MakePlanarJoint(nPlate, sTopPlatePort, -1, "Structure")
    AddJoint JointFactory.MakePrismaticJoint(nRod, "RodEnd1", nRod, "RodEnd2", 3574)
    AddJoint JointFactory.MakeRevoluteJoint(nRod, "RodEnd2", nPipeRing, "Side", 82)
    AddJoint JointFactory.MakeRigidJoint(nPipeRing, "Route", -1, "Route", 9444, , , dRodOffset + dSlideL / 2 - dSlideInset)

    AddJoint JointFactory.MakePlanarJoint(nRod2, "RodEnd1", nPlate, sBottomPlatePort)
    AddJoint JointFactory.MakePrismaticJoint(nRod2, "RodEnd1", nRod2, "RodEnd2", 3574)
    AddJoint JointFactory.MakeRevoluteJoint(nRod2, "RodEnd2", nPipeRing2, "Side", 82)
    AddJoint JointFactory.MakeRigidJoint(nPipeRing2, "Route", -1, "Route", 9444, , , dRodOffset - dSlideL / 2 + dSlideInset)

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
