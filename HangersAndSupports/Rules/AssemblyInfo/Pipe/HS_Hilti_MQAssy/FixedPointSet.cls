VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FixedPointSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright   (c) 2008, Intergraph Corporation.  All Rights reserved.
'
'   FixedPointSet.cls
'   ProgID:         HS_Hilti_MQ_Assy.FixedPointSet
'   Author:         JOGN
'   Createion Date: October 8, 2008
'   Description:    Fixed Point Set assembly for Hilti MQ library
'
'   Change History:
'   Date            who         change description
'   16-02-2010      VSP         Addition to S3D
'   04-04-2013      Ramya       CR 226479  FabRespons codelist should be removed from MQ Parts Codelist
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "FixedPointSet"

Implements IJHgrAssmInfo
Implements IJHgrSymbolBOMServices

Private numParts As Integer
Private nBase1 As Integer
Private nRod1 As Integer
Private nRod2 As Integer
Private nPipe1 As Integer
Private nPipeRing1 As Integer
Private nBPRod As Integer
Private nMFP As Integer
Private nMFP2 As Integer
Private nConnObj As Integer
Private nConnObj2 As Integer
Private nConnObj3 As Integer
Private nConnObj4 As Integer
Private nConnObj5 As Integer
Private nNut1 As Integer
Private nNut2 As Integer
Private nNut3 As Integer

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
Private lFabReq As Long
Private lFabType As Long
Private lConReq As Long
Private lConType As Long
Private lFabRespons As Long


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

    nBase1 = 1
    nPipe1 = 2
    nRod1 = 3
    nRod2 = 4
    nPipeRing1 = 5
    nBPRod = 6
    nMFP = 7
    nMFP2 = 8
    nConnObj = 9
    nConnObj2 = 10
    nNut1 = 11
    nNut2 = 12
    nNut3 = 13
    nConnObj3 = 14
    nConnObj4 = 15
    nConnObj5 = 16

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
    lFabReq = GetAttr("FabricationRequirement")
    lFabType = GetAttr("FabricationType")
    lConReq = GetAttr("ConstructionRequirement")
    lConType = GetAttr("ConstructionType")
    lFabRespons = GetAttr("FabricationResponsibility")

    numParts = 16
    ReDim PartClasses(numParts) As String
    PartClasses(nBase1) = sPlate
    PartClasses(nPipe1) = sPipe
    PartClasses(nRod1) = sRod
    PartClasses(nRod2) = sRod
    PartClasses(nPipeRing1) = sPipeRing
    PartClasses(nBPRod) = sBPRod
    PartClasses(nMFP) = sMFP
    PartClasses(nMFP2) = sMFP
    PartClasses(nConnObj) = "Log_Conn_Part_1"
    PartClasses(nConnObj2) = "Log_Conn_Part_1"
    PartClasses(nNut1) = sNut
    PartClasses(nNut2) = sNut
    PartClasses(nNut3) = sNut
    PartClasses(nConnObj3) = "Log_Conn_Part_1"
    PartClasses(nConnObj4) = "Log_Conn_Part_1"
    PartClasses(nConnObj5) = "Log_Conn_Part_1"

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

    ' Create a Joint Factory
    Dim JointFactory As New HgrJointFactory
    ' Create an object to hold the Joint
    Dim AssemblyJoint As Object

    Dim dStructAngle As Double
    Dim bstructStraight As Boolean
    Dim dStructRouteLengthHorz As Double
    Dim dAngleStructlength As Double
    Dim dPipeAngleX As Double
    Dim dStructAngleX As Double
    Dim bIsParrel As Boolean
    Dim dClampThick3 As Double
    Dim dPipeND As Double
    Dim sUnits As String
    Dim spart As String
    Dim sMFPInterface As String
    Dim dMFPThickness As Double
    Dim bStructIsSloped As Boolean
    Dim bPipeIsSloped As Boolean
    Dim bStructRoutePar As Boolean
    Dim bOrient As Boolean

    dStructAngle = GetPortOrientationAngle(ORIENT_GLOBAL_Z, "Structure", HGRPORT_Z)
    dPipeAngleX = GetPortOrientationAngle(ORIENT_GLOBAL_X, "Route", HGRPORT_X)
    dStructAngleX = GetPortOrientationAngle(ORIENT_GLOBAL_X, "Structure", HGRPORT_X)

    If (Deg(dStructAngle) > 180 - 0.00001 And Deg(dStructAngle) < 180 + 0.00001) Or (Deg(dStructAngle) > 0 - 0.00001 And Deg(dStructAngle) < 0 + 0.00001) Then
        bstructStraight = True
    Else
        bstructStraight = False
    End If

    SetPartAttribute nRod1, "Length", dRodLength, "Double"

    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationRequirement", IJElements_PartOccCollection.Item(nPipeRing1), lFabReq
    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationType", IJElements_PartOccCollection.Item(nPipeRing1), lFabType
    myIJHgrInputConfigHlpr.SetAttributeValue "ConstructionRequirement", IJElements_PartOccCollection.Item(nPipeRing1), lConReq
    myIJHgrInputConfigHlpr.SetAttributeValue "ConstructionType", IJElements_PartOccCollection.Item(nPipeRing1), lConType
    myIJHgrInputConfigHlpr.SetAttributeValue "FabricationResponsibility", IJElements_PartOccCollection.Item(nPipeRing1), lFabRespons


    bStructIsSloped = Hilti_StructIsSloped()
    bPipeIsSloped = Hilti_PipeIsSloped()

    GetNOMPipeDiaAndUnits 1, dPipeND, sUnits
    spart = Hilti_GetPartNumber(sPipeRing, dPipeND)

    dClampThick3 = Hilti_MultipleInterfaceDataLookUp("Thickness3", "JUAhsThickness3", "JCUHASMQ_LightFixPointMGALV", "PartNumber", "JDPart", "'" & spart & "'")
    dMFPThickness = Hilti_MultipleInterfaceDataLookUp("Thickness1", "JUAhsThickness1", "JCUHASMQ_BracingSetHDG", "PartNumber", "JDPart", "'" & sMFP & "'")


    Dim dBase1toStructConfig As Double
    Dim dBase1toClampConfig As Double
    Dim dBase1toStructPointOnConfig As Double
    Dim dClamptoPipeConfig As Double

    'CONFIG INDEXES
    '==============
    If Hilti_GetSupportingTypes(1) <> "Slab" Then
        'FOR STEEL
        dBase1toStructConfig = 36
        dBase1toClampConfig = 46
    Else
        'FOR SLAB
        If Hilti_RouteRunParallelSlopedSlab() Then
            'FOR SLAB SLOPED PARALLEL TO ROUTE
            dBase1toStructConfig = 36
            dBase1toClampConfig = 110
        Else
            'FOR SLAB SLOPED ACCROSS ROUTE
            dBase1toStructConfig = 53
            dBase1toClampConfig = 110
        End If
    End If

    If bPipeIsSloped And Hilti_GetSupportingTypes(1) <> "Slab" Then
        'FOR STEEL WITH SLOPED PIPE
        dBase1toStructPointOnConfig = 5
        dClamptoPipeConfig = 73
    ElseIf Hilti_GetSupportingTypes(1) = "Slab" And Hilti_RouteRunParallelSlopedSlab() Then
        'FOR SLAB SLOPED PARALLEL TO PIPE
        dBase1toStructPointOnConfig = 6
        dClamptoPipeConfig = 73
    Else
        'FOR THE REST
        dBase1toStructPointOnConfig = 4
        dClamptoPipeConfig = 73
    End If

    'DAS JOINTS
    '==========
    AddJoint JointFactory.MakePlanarJoint(nBase1, "Top", -1, "Structure", dBase1toStructConfig)
    AddJoint JointFactory.MakePlanarJoint(nBase1, "Top", nPipeRing1, "Route", dBase1toClampConfig)
    AddJoint JointFactory.MakePointOnJoint(nBase1, "Top", -1, "Structure", dBase1toStructPointOnConfig)
    AddJoint JointFactory.MakeSphericalJoint(nBase1, "Bottom", nRod1, "RodEnd1")
    AddJoint JointFactory.MakePrismaticJoint(nRod1, "RodEnd1", nRod1, "RodEnd2", 3574)
    AddJoint JointFactory.MakeVerticalJoint(nRod1, "RodEnd2", 67)
    AddJoint JointFactory.MakeRevoluteJoint(nRod1, "RodEnd2", nPipeRing1, "Side", 82)
    AddJoint JointFactory.MakeCylindricalJoint(nPipeRing1, "Route", -1, "Route", dClamptoPipeConfig)

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

    Set IJHgrAssmInfo_GetRouteConnections = ConnectPartToRouteOrStruct(nPipeRing1, 1)   'partindex, routeindex

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Function IJHgrAssmInfo_GetStructConnections(ByVal pDispInputConfigHlpr As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetStructConnections"

    On Error GoTo ErrorHandler

    Set IJHgrAssmInfo_GetStructConnections = ConnectPartToRouteOrStruct(nBase1, 1)

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


