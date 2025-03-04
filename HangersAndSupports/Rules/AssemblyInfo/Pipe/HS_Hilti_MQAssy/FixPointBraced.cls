VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FixPointBraced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright   (c) 2008, Intergraph Corporation.  All Rights reserved.
'
'   FixPointBraced.cls
'   ProgID:         HS_Hilti_MQAssy.FixPointBraced
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

Private Const MODULE = "FixPointBraced"

Implements IJHgrAssmInfo
Implements IJHgrSymbolBOMServices

Private numParts As Integer
Private nPlate As Integer
Private nRod1 As Integer
Private nRod2 As Integer
Private nPipe1 As Integer
Private nPipeRing As Integer
Private nBPRod As Integer
Private nMFP As Integer
Private nMFP2 As Integer
Private nNut1 As Integer
Private nNut2 As Integer
Private nNut3 As Integer
Private nConnObj As Integer
Private nMFP3 As Integer
Private nMFP4 As Integer
Private nNut4 As Integer
Private nNut5 As Integer
Private nRod3 As Integer
Private nConnObj2 As Integer
Private nConnObj3 As Integer
Private nConnObj4 As Integer

Private sPlate As String
Private sRod As String
Private sRod2 As String
Private sPipe As String
Private sMFP As String
Private sBPRod As String
Private sPipeRing As String
Private dH As Double
Private dRodLength As Double
Private sNut As String
Private sPartNum As String
Private lFabReq As Long
Private lFabType As Long
Private lConReq As Long
Private lConType As Long
Private lFabRespons As Long
Private dNumBrace As Double



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
    nPipe1 = 2
    nRod1 = 3
    nRod2 = 4
    nPipeRing = 5
    nBPRod = 6
    nMFP = 7
    nMFP2 = 8
    nNut1 = 9
    nNut2 = 10
    nNut3 = 11
    nConnObj = 12
    nConnObj3 = 13

    nMFP3 = 14
    nMFP4 = 15
    nNut4 = 16
    nNut5 = 17
    nRod3 = 18
    nConnObj2 = 19
    nConnObj4 = 20

    Initialize pDispInputConfigHlpr
    sPlate = GetAttr("Plate")
    sRod = GetAttr("Rod")
    sRod2 = GetAttr("Rod")
    sPipe = GetAttr("Pipe")
    sPipeRing = GetAttr("PipeRing")
    sMFP = GetAttr("MFP")
    sBPRod = GetAttr("BPRod")
    dH = GetAttr("H")
    dRodLength = GetAttr("RodLength")
    sNut = GetAttr("Nut")
    sPartNum = GetAttr("PartNum")
    dNumBrace = GetAttr("NumBrace")
    lFabReq = GetAttr("FabricationRequirement")
    lFabType = GetAttr("FabricationType")
    lConReq = GetAttr("ConstructionRequirement")
    lConType = GetAttr("ConstructionType")
    lFabRespons = GetAttr("FabricationResponsibility")

    Hilti_RouteSizeIsValid sPartNum
    
    If dNumBrace = -1 Then
        dNumBrace = dNumBrace * -1
    End If

    If dNumBrace = 1 Then
        numParts = 13
    Else
        numParts = 20
    End If
    ReDim PartClasses(numParts) As String
    PartClasses(nPlate) = sPlate
    PartClasses(nPipe1) = sPipe
    PartClasses(nRod1) = sRod
    PartClasses(nRod2) = sRod
    PartClasses(nPipeRing) = sPipeRing
    PartClasses(nBPRod) = sBPRod
    PartClasses(nMFP) = sMFP
    PartClasses(nMFP2) = sMFP
    PartClasses(nNut1) = sNut
    PartClasses(nNut2) = sNut
    PartClasses(nNut3) = sNut
    PartClasses(nConnObj) = "Log_Conn_Part_1"
    PartClasses(nConnObj3) = "Log_Conn_Part_1"

    If dNumBrace = 2 Then
        PartClasses(nMFP3) = sMFP
        PartClasses(nMFP4) = sMFP
        PartClasses(nNut4) = sNut
        PartClasses(nNut5) = sNut
        PartClasses(nRod3) = sRod
        PartClasses(nConnObj2) = "Log_Conn_Part_1"
        PartClasses(nConnObj4) = "Log_Conn_Part_1"
    End If

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

    Dim dClampThick3 As Double
    Dim spart As String
    Dim dMFPThickness As Double
    Dim sTopPlatePort As String
    Dim sBottomPlatePort As String

    Dim dNDFrom As Double

    SetPartAttribute nRod1, "Length", dRodLength, "Double"

    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationRequirement", IJElements_PartOccCollection.Item(nPipeRing), lFabReq
    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationType", IJElements_PartOccCollection.Item(nPipeRing), lFabType
    myIJHgrInputConfigHlpr.SetAttributeValue "ConstructionRequirement", IJElements_PartOccCollection.Item(nPipeRing), lConReq
    myIJHgrInputConfigHlpr.SetAttributeValue "ConstructionType", IJElements_PartOccCollection.Item(nPipeRing), lConType
    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationResponsibility", IJElements_PartOccCollection.Item(nPipeRing), lFabRespons

    Dim dBase1toStructConfig As Double
    Dim dBase1toClampConfig As Double
    Dim dBase1toStructPointOnConfig As Double
    Dim dClamptoPipeConfig As Double
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
    Else
        
    End If

    Dim dTempPipeND As Double
    dTempPipeND = GetNOMPipeDiaByUnit(1, NPD_INCH)

    spart = Hilti_GetPartNumber(sPipeRing, dTempPipeND)

    dClampThick3 = Hilti_MultipleInterfaceDataLookUp("Thickness3", "JUAhsThickness3", "JCUHASMQ_LightFixPointMGALV", "PartNumber", "JDPart", "'" & spart & "'")
    dMFPThickness = Hilti_MultipleInterfaceDataLookUp("Thickness1", "JUAhsThickness1", "JCUHASMQ_BracingSetHDG", "PartNumber", "JDPart", "'" & sMFP & "'")

    If Hilti_StructIsSloped() Then
        PF_EventHandler "Assembly can not be placed to sloped Slab/Structure.", Err, MODULE, METHOD, False
    End If

    'Set config idx for Pipe for Route Connection value
    Dim iRingPipeConfigIdx As Integer, iMFPStructIdx As Integer
    If GetRouteConnectionValue = 1 Then
        iRingPipeConfigIdx = 1252
        Else
        iRingPipeConfigIdx = 9444
        End If

    AddJoint JointFactory.MakePlanarJoint(nPlate, sTopPlatePort, -1, "Structure", iPlateStructConfigIdx)
    AddJoint JointFactory.MakeRigidJoint(nPlate, sBottomPlatePort, nPipe1, "RodEnd1", iPlateRodConfigIdx)
    AddJoint JointFactory.MakePrismaticJoint(nPipe1, "RodEnd1", nPipe1, "RodEnd2", 3574)
    AddJoint JointFactory.MakeRigidJoint(nPipe1, "RodEnd2", nBPRod, "Rod", 9444)
    AddJoint JointFactory.MakeRigidJoint(nBPRod, "Rod2", nRod1, "RodEnd1", 9444)
    AddJoint JointFactory.MakeRevoluteJoint(nRod1, "RodEnd2", nPipeRing, "Side", 82)
    AddJoint JointFactory.MakeRigidJoint(nPipeRing, "Route", -1, "Route", iRingPipeConfigIdx)

'    ################################################################################################
'     The Following Code will generate joints for the first brace
    Dim iMFPConnConfigIdx As Integer
    If oICH.IsPlaceByStructure = False Then
        iMFPConnConfigIdx = 1188
    Else
        iMFPConnConfigIdx = 2212
    End If
    AddJoint JointFactory.MakeRevoluteJoint(nConnObj, "Connection", nPipeRing, "Side", 82)
    AddJoint JointFactory.MakeRigidJoint(nConnObj3, "Connection", nConnObj, "Connection", 9444, -dMFPThickness / 2 - dClampThick3)
    AddJoint JointFactory.MakeRigidJoint(nMFP, "Rod", nConnObj3, "Connection", iMFPConnConfigIdx)    '91)

    If dNumBrace = 1 Then
        AddJoint JointFactory.MakeRigidJoint(nMFP, "Rod", nNut1, "Bottom", 9380, -dMFPThickness / 2)
    End If

    AddJoint JointFactory.MakeRigidJoint(nMFP, "Rod2", nNut2, "Top", 9444, -dMFPThickness)
    AddJoint JointFactory.MakeRigidJoint(nNut2, "Bottom", nRod2, "RodEnd1", 9444, -dMFPThickness)
    AddJoint JointFactory.MakePrismaticJoint(nRod2, "RodEnd1", nRod2, "RodEnd2", 3574)
    AddJoint JointFactory.MakeRigidJoint(nMFP2, "Rod2", nNut3, "Bottom", 9380, -dMFPThickness)
    AddJoint JointFactory.MakeRigidJoint(nNut3, "Top", nRod2, "RodEnd2", 9380, dMFPThickness)
    AddJoint JointFactory.MakePlanarJoint(nMFP2, "Rod", -1, "Structure", 36, dMFPThickness / 2)
'
'    '################################################################################################
'    ' The Following Code will generate joints for the Second brace if needed
    If dNumBrace = 2 Then

        AddJoint JointFactory.MakeRigidJoint(nMFP, "Rod", nNut1, "Bottom", 9380, -dMFPThickness * 1.5)

        AddJoint JointFactory.MakePlanarJoint(nConnObj4, "Connection", -1, "Structure", 100, -dMFPThickness / 2)

        AddJoint JointFactory.MakePlanarJoint(nConnObj4, "Connection", nMFP2, "Rod", 45)

        AddJoint JointFactory.MakeRigidJoint(nMFP, "Rod", nMFP3, "Rod", 1252, -dMFPThickness)
        Rigid nMFP3, "Rod2", nNut4, "Bottom", 0, 0, -dMFPThickness, 0, 180
        Rigid nNut4, "Top", nRod3, "RodEnd1", 0, 0, dMFPThickness, 0, 180
        AddJoint JointFactory.MakePrismaticJoint(nRod3, "RodEnd1", nRod3, "RodEnd2", 3574)
        Rigid nMFP4, "Rod2", nNut5, "Bottom", 0, 0, -dMFPThickness, 0, 180
        Rigid nNut5, "Top", nRod3, "RodEnd2", 0, 0, dMFPThickness, 0, 180
        AddJoint JointFactory.MakePointOnJoint(nMFP4, "Rod", nConnObj4, "Connection", 1)
    End If
    '################################################################################################
    ' Return the collection of Joints
    Set IJHgrAssmInfo_GetAssemblyJoints = GetJoints

    Class_Terminate

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Function IJHgrAssmInfo_GetMaxRouteConnectionValue() As Long

    IJHgrAssmInfo_GetMaxRouteConnectionValue = 2

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

    bstrBOMDesc = "Hilti"


    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

