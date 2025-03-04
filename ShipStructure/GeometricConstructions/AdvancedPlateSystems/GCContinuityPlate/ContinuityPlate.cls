VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContinuityPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum Errors
    MISSING_MEMBER_PART = 1
    MISSING_PLATE_SYSTEM = 2
    MISSING_SIDE_PLATE1 = 3
    MISSING_SIDE_PLATE2 = 4
    MISSING_REFERENCE_MEMBER_PART = 5
    DUPLICATE_INPUTS = 6
    CANNOT_RETRIEVE_TRIMMING_PLANE = 7
    CANNOT_BE_BOUNDED_BY_PLATE_SYSTEM = 8
End Enum

Private Const sMEMBER_PART As String = "MemberPart"
Private Const sPLATE_SYSTEM As String = "PlateSystem"
Private Const sREFERENCE_MEMBER_PART As String = "ReferenceMemberPart"
Private Const sSIDE_PLATE1 As String = "SidePlate1"
Private Const sSIDE_PLATE2 As String = "SidePlate2"
Private Const sEXTENSION As String = "Extension"
Private Const sLOCATION As String = "Location"

Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGCMigrate
Implements IJContinuityPlate
Implements IJGCToDoDelegate
Implements IJGCConnectedElementSC

Sub BreakPoint()
    Dim i As Integer: Let i = 0
End Sub
Private Property Get Source() As String
    Let Source = "GCContinuityPlate.ContinuityPlate"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property
Private Sub Class_Initialize()
    Call DebugStart(Source)

'    Let m_lCountOfDebugSupport = m_lCountOfDebugSupport + 1
'    Set m_pDebugSupport(m_lCountOfDebugSupport) = New DebugSupport
'    Let m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_SOURCE = "ContinuityPlate1 (" + CStr(m_lCountOfDebugSupport) + ")"
'    Call m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_MSG(">> ContinuityPlate1")
End Sub
Private Sub Class_Terminate()
'    Call m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_MSG("<< ContinuityPlate1")
'    Set m_pDebugSupport(m_lCountOfDebugSupport) = Nothing
'    Let m_lCountOfDebugSupport = m_lCountOfDebugSupport - 1

    Call DebugStop(Source)
End Sub

Private Sub IJGCConnectedElementSC_PostDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)

End Sub

Private Sub IJGCConnectedElementSC_PreDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)
    If pGC.ControlledInputs("AdvancedPlateSystem").Count = 1 Then

        ' retrieve advanced plate system
        Dim oPlateSystemAdvanced As Object: Set oPlateSystemAdvanced = pGC.ControlledInputs("AdvancedPlateSystem")(1)

        ' retrieve info about the incoming members
        Dim iMember As Integer
        Dim oInputMember As Object
        Dim pElementsOfPrimaryMembers As IJElements: Set pElementsOfPrimaryMembers = pGC.Inputs(sMEMBER_PART)
        Dim iCountOfPrimaryMembers As Integer: iCountOfPrimaryMembers = pElementsOfPrimaryMembers.Count

        ' retrieve info about the secondary members
        Dim pElementsOfSecondaryMembers As IJElements: Set pElementsOfSecondaryMembers = pGC.Inputs(sREFERENCE_MEMBER_PART)
        Dim iCountOfSecondaryMembers As Integer: iCountOfSecondaryMembers = pElementsOfSecondaryMembers.Count

        ' retrieve the node
        Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGC
        Dim pPositionOfNode As IJDPosition: Set pPositionOfNode = Position_FromPoint(pGeometricConstructionMacro.Outputs("Node")(1))

        Dim bAPSTransformed As Boolean
        Dim bInputMemberTransformed As Boolean
        bAPSTransformed = pInfo.IsDesignObjectInTransformSet(oPlateSystemAdvanced)
        If bAPSTransformed Then
            'APS transformed, unbound members that are not transformed
            For iMember = 1 To iCountOfPrimaryMembers
                Set oInputMember = pElementsOfPrimaryMembers(iMember)
                bInputMemberTransformed = pInfo.IsDesignObjectInTransformSet(oInputMember)
                If Not bInputMemberTransformed Then
                    'unbound member
                    Call Member_UnBoundFromEdgeOfAPS(oInputMember, pGC, oPlateSystemAdvanced, pPositionOfNode)
                End If
            Next
            For iMember = 1 To iCountOfSecondaryMembers
                Set oInputMember = pElementsOfSecondaryMembers(iMember)
                bInputMemberTransformed = pInfo.IsDesignObjectInTransformSet(oInputMember)
                If Not bInputMemberTransformed Then
                    'unbound member
                    Call Member_BoundByOrUnBoundFromFaceOfAPS(oInputMember, False, oPlateSystemAdvanced, pPositionOfNode)
                End If
            Next
        End If
    End If

End Sub

'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Initialize"))

    ' located input
    Call pGeometricConstructionDefinition.AddInput(sMEMBER_PART, "Select a Member", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput(sPLATE_SYSTEM, "Select a limiting Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "")
    ' limit next 2 inputs to be APS
    Call pGeometricConstructionDefinition.AddInput(sSIDE_PLATE1, "Select the back SidePlate (APS)", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput(sSIDE_PLATE2, "Select the front SidePlate (APS)", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput(sREFERENCE_MEMBER_PART, "Select a Member as reference for the positioning", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    ' controlled inputs
    Call pGeometricConstructionDefinition.AddControlledInput("AdvancedPlateSystem", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort2")
    Call pGeometricConstructionDefinition.AddControlledInput("PlateSystem2", "")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort3")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort4")

    ' special controlled input to indicate that some PlateSystems of incoming Members have not been correctly un-bounded during the disconnect semantic,
    ' in the case, where the APS and some incoming Members are simultaneously deleted.
    ' this is due to the fact that FrameConnections depending from a deleted Member are first deleted, before being later recreated.
    ' our disconnect semantic executes, after those FrameConnections get deleted and before they get re-created.
    ' we have to wait the triggering of our compute semantic to have access again to the re-created ones.
    Call pGeometricConstructionDefinition.AddControlledInput("UnBoundedPlateSystems", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    ' parameters
    Call pGeometricConstructionDefinition.AddParameter(sEXTENSION, sEXTENSION, GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    
    ' code listed parameters
    Call pGeometricConstructionDefinition.AddParameter(sLOCATION, sLOCATION, GCCodeList, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue(sLOCATION, "Near", 1)
    Call pGeometricConstructionDefinition.AddParameterValue(sLOCATION, "Far", 2)
    Call pGeometricConstructionDefinition.AddParameterValue(sLOCATION, "Centered", 3)

    ' error codes
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MEMBER_PART, "MissingMemberPart", "Missing the member part")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_PLATE_SYSTEM, "MissingPlateSystem", "Missing the plate system")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_SIDE_PLATE1, "MissingSidePlate1", "Missing the side plate (APS) #1")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_SIDE_PLATE2, "MissingSidePlate2", "Missing the side plate (APS) #2")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_REFERENCE_MEMBER_PART, "MissingReferenceMemberPart", "Missing the reference member part")
    Call pGeometricConstructionDefinition.AddErrorValue(DUPLICATE_INPUTS, "DuplicateInputs", "2 inputs are the same")
    Call pGeometricConstructionDefinition.AddErrorValue(CANNOT_RETRIEVE_TRIMMING_PLANE, "CannotRetrieveTrimmingPlane", "Cannot retrieve trimming plane along MemberPart from side plate (APS) #1")
    Call pGeometricConstructionDefinition.AddErrorValue(CANNOT_BE_BOUNDED_BY_PLATE_SYSTEM, "CannotBeBoundedByPlateSystem", "The unbounded geometry of the limiting Web or PlateSystem is not big enough to bound the ContinuityPlate (enhance the 'Length Extension' attribute of the Built-Up)")
    
    ' outputs
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Support")
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypePoint3d, "Node")
    Call pGeometricConstructionDefinition.AddOutput(GCLocalCoordinateSystem, "CoordinateSystem")
    
    Call DebugOut
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    On Error GoTo ErrorHandler: Call BreakPoint
    
    ' use GC as GCMacro
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGeometricConstruction
    
    ' continue the work partially done during the disconnect semantic
    If pGeometricConstruction.ControlledInputs("UnBoundedPlateSystems").Count > 0 Then
        On Error Resume Next
        Call Members_UnboundOnDelete(pGeometricConstruction, pPOM, Nothing)
        On Error GoTo 0
        
        ' delete myself, because the APS is deleted and we do not want any more the compute semantic to fire
        Call Object_Delete(pGeometricConstruction)
        Exit Sub
    End If
    
    ' get factories
    Dim pGCFactory As IJGeometricConstructionEntitiesFactory: Set pGCFactory = New GeometricConstructionEntitiesFactory
        
    ' check for errors on inputs only in compute mode
    If True Then
        ' check mandatory inputs
        If pGeometricConstruction.Inputs(sMEMBER_PART).Count = 0 Then Err.Raise MISSING_MEMBER_PART
        If pGeometricConstruction.Inputs(sPLATE_SYSTEM).Count = 0 Then Err.Raise MISSING_PLATE_SYSTEM
        If pGeometricConstruction.Inputs(sSIDE_PLATE1).Count = 0 Then Err.Raise MISSING_SIDE_PLATE1
        If pGeometricConstruction.Inputs(sSIDE_PLATE2).Count = 0 Then Err.Raise MISSING_SIDE_PLATE2
        If pGeometricConstruction.Inputs(sREFERENCE_MEMBER_PART).Count = 0 Then Err.Raise MISSING_REFERENCE_MEMBER_PART
        
        ' check duplicate inputs
        Dim pElements As IJElements: Set pElements = New JObjectCollection
        Call pElements.Add(pGeometricConstruction.Input(sMEMBER_PART))
        If pElements.Contains(pGeometricConstruction.Input(sPLATE_SYSTEM)) Then Err.Raise DUPLICATE_INPUTS
        Call pElements.Add(pGeometricConstruction.Input(sPLATE_SYSTEM))
        If pElements.Contains(pGeometricConstruction.Input(sSIDE_PLATE1)) Then Err.Raise DUPLICATE_INPUTS
        Call pElements.Add(pGeometricConstruction.Input(sSIDE_PLATE1))
        If pElements.Contains(pGeometricConstruction.Input(sSIDE_PLATE2)) Then Err.Raise DUPLICATE_INPUTS
        Call pElements.Add(pGeometricConstruction.Input(sSIDE_PLATE2))
        If pElements.Contains(pGeometricConstruction.Input(sREFERENCE_MEMBER_PART)) Then Err.Raise DUPLICATE_INPUTS
    End If
        
    ' retrieve info about the input member
    Dim pElementsOfMembers As IJElements: Set pElementsOfMembers = pGeometricConstruction.Inputs(sMEMBER_PART)
    Dim iCountOfMembers As Integer:  Let iCountOfMembers = pElementsOfMembers.Count
    Dim sKeysOfMembers() As String: If iCountOfMembers > 0 Then Let sKeysOfMembers = Elements_GetKeys(pElementsOfMembers)
        
    ' retrieve info about boundaries
    Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
    Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
    Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)
        
    ' retrieve info about the PlateSystem
    Dim oPlateSystem As Object: Set oPlateSystem = pGeometricConstruction.Input(sPLATE_SYSTEM)
    
    ' retrieve info about the ReferenceMember
    Dim oReferenceMember As Object: Set oReferenceMember = pGeometricConstruction.Input(sREFERENCE_MEMBER_PART)
    
    ' retrieve SidePlates
    Dim oSidePlate1 As Object: Set oSidePlate1 = pGeometricConstruction.Input(sSIDE_PLATE1)
    Dim oSidePlate2 As Object: Set oSidePlate2 = pGeometricConstruction.Input(sSIDE_PLATE2)
    
    ' retrieve trimming plane and position of node
    Dim oTrimmingPlane As Object
    Dim pPositionOfNode As IJDPosition
    
    Call GetTrimmingPlaneAndNodePosition(pPOM, pElementsOfMembers(1), oSidePlate1, oSidePlate2, oTrimmingPlane, pPositionOfNode)
    If oTrimmingPlane Is Nothing Then Err.Raise CANNOT_RETRIEVE_TRIMMING_PLANE
   
    ' compute the coordinate system
    Dim pCoordinateSystem As IJDCoordinateSystem: Set pCoordinateSystem = ComputeCS(pGeometricConstruction, pPOM, pGCFactory, _
                                                                                    MemberPart_GetLine(pElementsOfMembers(1), pPOM), _
                                                                                    MemberPart_GetLine(oReferenceMember, pPOM))

    ' un-bound removed members
    If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        Call Members_UnboundFromEdgeOfAPS(pElementsOfMembers, iCountOfMembers, sKeysOfMembers, _
                                          pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                          pGeometricConstruction, 1, pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1), pPositionOfNode)
    End If
    
    ' check if the PlateSystem is a ContinuityPlate
    Dim bIsContinuityPlate As Boolean: Let bIsContinuityPlate = False
    If True Then
        Dim oCreateInstance As New CreateInstance
        Dim oFilters As Object
        Set oFilters = oCreateInstance.SP3DCreateObject("GCFilters.Filters")
        Dim iIsContinuityPlate As Integer: Let iIsContinuityPlate = oFilters.IsContinuityPlate(oPlateSystem)
        If iIsContinuityPlate = 1 Then Let bIsContinuityPlate = True
    End If
    
    ' compute outputs
    Call ComputeOutputs(pGeometricConstruction, pPOM, pGCFactory, _
                        pCoordinateSystem, _
                        pElementsOfMembers(1), sKeysOfMembers(1), _
                        oPlateSystem, bIsContinuityPlate, _
                        oReferenceMember, _
                        oSidePlate1, _
                        oSidePlate2, _
                        oTrimmingPlane, _
                        pPositionOfNode)
        
    If iCountOfBoundaries > 0 Then
        Call Keys_RemoveKey(sKeysOfBoundaries, "TR1." + sKeysOfMembers(1))
        If bIsContinuityPlate Then
            Call Keys_RemoveKey(sKeysOfBoundaries, "PlateSystem2")
        Else
            Call Keys_RemoveKey(sKeysOfBoundaries, "FacePort2")
        End If
        Call Keys_RemoveKey(sKeysOfBoundaries, "FacePort3")
        Call Keys_RemoveKey(sKeysOfBoundaries, "FacePort4")
    End If
    
    ' trim members
    If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        ' get the key of the APS
        Dim sKey As String: Let sKey = pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").GetKey(pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1))
        
        ' if the key is not set to "NoBound" then do the bound (this means that the APS has already computed its geometry and edge ports can be extracted)
        If Not sKey = "NotBounded" And Not sKey = "JustMigrated" Then
            Call Members_BoundByOrUnboundFromEdgeOfAPS(pElementsOfMembers, iCountOfMembers, sKeysOfMembers, _
                                                       pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                                       pGeometricConstruction, 1, pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1), pPositionOfNode)
        End If
    
        If sKey = "JustMigrated" Then
            Call Elements_RenameKeyOfElement(pGeometricConstruction.ControlledInputs("AdvancedPlateSystem"), "JustMigrated", "")
            ' force a recompute of the GCMacro
            Call pGeometricConstruction.UpdateGeometricConstructionAE
        End If
    End If
    
    ' remove unneeded boundaries
    If iCountOfBoundaries > 0 Then Call Elements_RemoveUnneeded(pElementsOfBoundaries, sKeysOfBoundaries)
    
    Call DebugOut
    Exit Sub
ErrorHandler:
    Call GCProcessError(pGeometricConstruction, , Err.Number)
End Sub
'
' implementation of the IJContinuityPlate interface
'
Private Function IJContinuityPlate_Create(ByVal pPOM As IJDPOM, _
                                          ByVal oMemberPart As Object, _
                                          ByVal oPlateSystem As Object, _
                                          ByVal oReferenceMemberPart As Object, _
                                          ByVal oSidePlate1 As Object, _
                                          ByVal oSidePlate2 As Object, _
                                          Optional ByVal dExtension As Double = 0#, _
                                          Optional ByVal eContinuityPlateLocation As eContinuityPlateLocation = 1) As Object
    ' create factory
    Dim pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory
    Set pGeometricConstructionFactory = New GeometricConstructionEntitiesFactory

    ' create geometric construction
    Dim pGeometricConstruction As IJGeometricConstruction
    Set pGeometricConstruction = pGeometricConstructionFactory.CreateEntity("ContinuityPlate", pPOM)
    
    ' provide inputs and parameters
    Call pGeometricConstruction.Inputs(sMEMBER_PART).Add(oMemberPart, "1")
    Call pGeometricConstruction.Inputs(sPLATE_SYSTEM).Add(oPlateSystem, "1")
    Call pGeometricConstruction.Inputs(sREFERENCE_MEMBER_PART).Add(oReferenceMemberPart, "1")
    Call pGeometricConstruction.Inputs(sSIDE_PLATE1).Add(oSidePlate1, "1")
    Call pGeometricConstruction.Inputs(sSIDE_PLATE2).Add(oSidePlate2, "1")
    pGeometricConstruction.Parameter(sEXTENSION) = dExtension
    pGeometricConstruction.Parameter(sLOCATION) = eContinuityPlateLocation
    
    ' return geometric construction
    Set IJContinuityPlate_Create = pGeometricConstruction
End Function
'
' implementation of the IJGCSemanticConnection interface
'
Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal oRelationship As Object)
    Call DebugIn(Method("IJGCSemanticConnection_PostConnectionAdded"))
    
    Dim pRelationship As IJDRelationship: Set pRelationship = oRelationship
    Call DebugInput("Origin", pRelationship.Origin)
    Call DebugInput("Destination", pRelationship.Destination)
    Call DebugInput("Name", pRelationship.Name)
    
    Call DebugOut
End Sub
Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
    On Error GoTo ErrorHandler: Call BreakPoint
    Call DebugIn(Method("IJGCSemanticConnection_PreConnectionRemoved"))
    
    Dim pRelationship As IJDRelationship: Set pRelationship = oRelationship
    Call DebugInput("Origin", pRelationship.Origin)
    Call DebugInput("Destination", pRelationship.Destination)
    Call DebugInput("Name", pRelationship.Name)
    Call DebugInput("IsOriginDeleted", bIsOriginDeleted)
    Call DebugInput("IsDestinationDeleted", bIsDestinationDeleted)
    
    Dim pGeometricConstruction As IJGeometricConstruction: Set pGeometricConstruction = pRelationship.Origin
    
    If Mid(pRelationship.Name, 1, Len("AdvancedPlateSystem")) = "AdvancedPlateSystem" And bIsDestinationDeleted Then
        ' the APS is deleting
        ' unbound remaining connected Members
        
        ' get POM for debug
        Dim pPOM As IJDPOM: ' Set pPOM = GetPOMFromObject(pGeometricConstruction)
    
        On Error Resume Next ' avoid failure to not prevent the APS to get deleted
        Call Members_UnboundOnDelete(pGeometricConstruction, pPOM, pRelationship.Destination)
        On Error GoTo ErrorHandler
    
        ' delete the GCMacro itself, if no more processing is expected from the compute semantic
        If pGeometricConstruction.ControlledInputs("UnBoundedPlateSystems").Count = 0 Then Call Object_Delete(pGeometricConstruction)
    ElseIf Mid(pRelationship.Name, 1, Len(sMEMBER_PART)) = sMEMBER_PART And bIsDestinationDeleted Then
        ' an input member is deleting, disconnect the corresponding ports
        Dim sKey As String: Let sKey = Mid(pRelationship.Name, Len(sMEMBER_PART) + 2)
        pGeometricConstruction.ControlledInputs("AxisPort1").Remove (sKey)
        pGeometricConstruction.ControlledInputs("FacePort1").Remove (sKey)
    End If
    
    Call DebugOut
    Exit Sub
ErrorHandler:
End Sub
'
' private APIs
'
Private Sub Members_UnboundOnDelete(pGeometricConstructionMacro As IJGeometricConstructionMacro, pPOM As IJDPOM, oAdvancedPlateSystem As Object)
    Call DebugIn(Method("Members_UnboundOnDelete"))

    ' use GCMacro as GC
    Dim pGeometricConstruction As IJGeometricConstruction: Set pGeometricConstruction = pGeometricConstructionMacro
   
    ' get factory
    Dim pGCFactory As IJGeometricConstructionEntitiesFactory
    Set pGCFactory = New GeometricConstructionEntitiesFactory
        
    ' retrieve info about the input member
    Dim pElementsOfMembers As IJElements: Set pElementsOfMembers = Elements_GetValidElements(pGeometricConstruction.Inputs(sMEMBER_PART))
    Dim iCountOfMembers As Integer:  Let iCountOfMembers = pElementsOfMembers.Count
    Dim sKeysOfMembers() As String: If iCountOfMembers > 0 Then Let sKeysOfMembers = Elements_GetDummyKeys(pElementsOfMembers)
    
    ' retrieve info about the ReferenceMember
    Dim oReferenceMember As Object
    If True Then
        Dim pElementsOfValidElements As IJElements:
        Set pElementsOfValidElements = Elements_GetValidElements(pGeometricConstruction.Inputs(sREFERENCE_MEMBER_PART))
        If pElementsOfValidElements.Count = 1 Then Set oReferenceMember = pElementsOfValidElements(1)
    End If
            
    If iCountOfMembers = 1 Then
        ' retrieve info about boundaries
        Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
        Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
        Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)
   
        ' compute the node
        Dim pPositionOfNode As IJDPosition
        Set pPositionOfNode = GeometricConstructionMacro_RetrievePositionOfNode(pGeometricConstructionMacro)
        
        If pPositionOfNode Is Nothing And Not oReferenceMember Is Nothing Then
            ' if something goes wrong, we will at least try to retrieve the node
            On Error Resume Next
            
            ' compute the coordinate system
            Dim pCoordinateSystem As IJDCoordinateSystem
            Set pCoordinateSystem = ComputeCS(pGeometricConstruction, pPOM, pGCFactory, _
                                              MemberPart_GetLine(pElementsOfMembers(1)), _
                                              MemberPart_GetLine(oReferenceMember))
            If Not pCoordinateSystem Is Nothing Then Set pPositionOfNode = pCoordinateSystem.Origin
            
            On Error GoTo 0
        End If
        
        
        ' un-bound removed members
        Call Members_UnboundFromEdgeOfAPS(pElementsOfMembers, iCountOfMembers, sKeysOfMembers, _
                                          pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                          pGeometricConstruction, 1, oAdvancedPlateSystem, pPositionOfNode)
                                          
    Call DebugOut
    End If
End Sub
Private Function ComputeCS(pGeometricConstruction As IJGeometricConstruction, pPOM As IJDPOM, pGCFactory As IJGeometricConstructionEntitiesFactory, _
                   pLineOfMember As IJLine, _
                   pLineReferenceMember As IJLine) As IJDCoordinateSystem
    Call DebugIn(Method("ComputeCS"))

'''    Dim oLineAxisPortExtractor0 As SP3DGeometricConstruction.GeometricConstruction
'''    Set oLineAxisPortExtractor0 = pGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "000-")
'''    oLineAxisPortExtractor0.Inputs(sMEMBER_PART).Add oReferenceMember
'''    oLineAxisPortExtractor0.Public = True
'''    oLineAxisPortExtractor0.Evaluate
'''
'''    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
'''    Set oLineAxisPortExtractor1 = pGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "000-")
'''    oLineAxisPortExtractor1.Inputs(sMEMBER_PART).Add oMember
'''    oLineAxisPortExtractor1.Public = True
'''    oLineAxisPortExtractor1.Evaluate

    Dim oCSByLines1 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines1 = pGCFactory.CreateEntity("CSByLines", pPOM, "001-CSByLines")
    oCSByLines1.Inputs("AxisLine1").Add pLineOfMember 'oLineAxisPortExtractor1
    oCSByLines1.Inputs("AxisLine2").Add pLineReferenceMember 'oLineAxisPortExtractor0
    oCSByLines1.Parameter("AxesRoles") = 2
    oCSByLines1.Parameter("CSOrientation") = 1
    oCSByLines1.Public = True
    oCSByLines1.Evaluate
    
    ' return result
    Set ComputeCS = oCSByLines1
    
    Call DebugOut
End Function
Private Sub ComputeOutputs(pGeometricConstruction As IJGeometricConstruction, pPOM As IJDPOM, pGCFactory As IJGeometricConstructionEntitiesFactory, _
                   pCoordinateSystem As IJDCoordinateSystem, _
                   oMember As Object, sKeyOfMember As String, _
                   oPlateSystem As Object, bIsContinuityPlate As Boolean, _
                   oReferenceMember As Object, _
                   oSidePlate1 As Object, _
                   oSidePlate2 As Object, _
                   oTrimmingPlane As Object, _
                   pPositionOfNode As IJDPosition)
    Call DebugIn(Method("ComputeOutputs"))
    
    Dim oLineAxisPortExtractor0 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor0 = pGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "002-LineAxisPortExtractor")
    oLineAxisPortExtractor0.Inputs(sMEMBER_PART).Add oReferenceMember
    oLineAxisPortExtractor0.Public = True
    oLineAxisPortExtractor0.Evaluate
    
    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = pGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "003-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs(sMEMBER_PART).Add oMember
    oLineAxisPortExtractor1.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("AxisPort1")
    oLineAxisPortExtractor1.Public = True
    oLineAxisPortExtractor1.Evaluate
    Call Elements_ReplaceElementsWithKey(pGeometricConstruction.ControlledInputs("AxisPort1"), oLineAxisPortExtractor1.ControlledInputs("Port"), sKeyOfMember)
    
    Dim oPointFromCS1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS1 = pGCFactory.CreateEntity("PointFromCS", pPOM, "004-PointFromCS")
    oPointFromCS1.Inputs("CoordinateSystem").Add pCoordinateSystem
    oPointFromCS1.Parameter("X") = 0#
    oPointFromCS1.Parameter("Y") = 0#
    oPointFromCS1.Parameter("Z") = 0#
    oPointFromCS1.Public = True
    oPointFromCS1.Evaluate
        
    Dim oFacePortExtractor2a As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor2a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "005-FacePortExtractor")
    oFacePortExtractor2a.Inputs("Connectable").Add oPlateSystem
'    If Not bIsContinuityPlate Then
'        oFacePortExtractor2a.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort2")
'    End If
    oFacePortExtractor2a.Parameter("GeometrySelector") = 4
    oFacePortExtractor2a.Parameter("FacesContext") = 1
    oFacePortExtractor2a.Parameter("LookingAxis") = 1
    oFacePortExtractor2a.Parameter("TrackFlag") = 1
    oFacePortExtractor2a.Parameter("Offset") = 0
    oFacePortExtractor2a.Public = True
    oFacePortExtractor2a.Evaluate
    If Not bIsContinuityPlate Then
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort2"), oFacePortExtractor2a.ControlledInputs("Port"))
        Call pGeometricConstruction.ControlledInputs("PlateSystem2").Clear
    Else
        Dim pElementsOfPlateSystem As IJElements: Set pElementsOfPlateSystem = New JObjectCollection
        Call pElementsOfPlateSystem.Add(oPlateSystem)
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("PlateSystem2"), pElementsOfPlateSystem)
        Call pGeometricConstruction.ControlledInputs("FacePort2").Clear
    End If
    
    Dim oFacePortExtractor3a As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor3a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "006-FacePortExtractor")
    oFacePortExtractor3a.Inputs("Connectable").Add oSidePlate1
    oFacePortExtractor3a.Parameter("GeometrySelector") = 2
    oFacePortExtractor3a.Parameter("FacesContext") = 1
    oFacePortExtractor3a.Parameter("LookingAxis") = 1
    oFacePortExtractor3a.Parameter("TrackFlag") = 1
    oFacePortExtractor3a.Parameter("Offset") = 0
    oFacePortExtractor3a.Public = True
    oFacePortExtractor3a.Evaluate
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort3"), oFacePortExtractor3a.ControlledInputs("Port"))
    
    Dim oFacePortExtractor4a As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor4a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "007-FacePortExtractor")
    oFacePortExtractor4a.Inputs("Connectable").Add oSidePlate2
    oFacePortExtractor4a.Parameter("GeometrySelector") = 2
    oFacePortExtractor4a.Parameter("FacesContext") = 1
    oFacePortExtractor4a.Parameter("LookingAxis") = 1
    oFacePortExtractor4a.Parameter("TrackFlag") = 1
    oFacePortExtractor4a.Parameter("Offset") = 0
    oFacePortExtractor4a.Public = True
    oFacePortExtractor4a.Evaluate
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort4"), oFacePortExtractor4a.ControlledInputs("Port"))
    
    If CInt(pGeometricConstruction.Parameter(sLOCATION)) <> 3 Then
        Dim oFacePortExtractor1a As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor1a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "008-FacePortExtractor")
        oFacePortExtractor1a.Inputs("Connectable").Add oMember
        oFacePortExtractor1a.Inputs("CoordinateSystem").Add pCoordinateSystem
        oFacePortExtractor1a.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort1")
        oFacePortExtractor1a.Parameter("GeometrySelector") = 4
        oFacePortExtractor1a.Parameter("FacesContext") = 2
        oFacePortExtractor1a.Parameter("LookingAxis") = 3
        If pGeometricConstruction.Parameter(sLOCATION) = 1 Then
            oFacePortExtractor1a.Parameter("TrackFlag") = 2
        Else
            oFacePortExtractor1a.Parameter("TrackFlag") = 1
        End If
        oFacePortExtractor1a.Parameter("Offset") = 0#
        oFacePortExtractor1a.Public = True
        oFacePortExtractor1a.Evaluate
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort1"), oFacePortExtractor1a.ControlledInputs("Port"), sKeyOfMember)
    Else
        Set oFacePortExtractor1a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "008a-FacePortExtractor")
        oFacePortExtractor1a.Inputs("Connectable").Add oMember
        oFacePortExtractor1a.Inputs("CoordinateSystem").Add pCoordinateSystem
        oFacePortExtractor1a.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort1")
        oFacePortExtractor1a.Parameter("GeometrySelector") = 4
        oFacePortExtractor1a.Parameter("FacesContext") = 2
        oFacePortExtractor1a.Parameter("LookingAxis") = 3
        oFacePortExtractor1a.Parameter("TrackFlag") = 1
        oFacePortExtractor1a.Parameter("Offset") = 0#
        oFacePortExtractor1a.Public = True
        oFacePortExtractor1a.Evaluate
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort1"), oFacePortExtractor1a.ControlledInputs("Port"), sKeyOfMember)
        
        Dim oFacePortExtractor1b As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor1b = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "008b-FacePortExtractor")
        oFacePortExtractor1b.Inputs("Connectable").Add oMember
        oFacePortExtractor1b.Inputs("CoordinateSystem").Add pCoordinateSystem
        'oFacePortExtractor1b.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort1")
        oFacePortExtractor1b.Parameter("GeometrySelector") = 4
        oFacePortExtractor1b.Parameter("FacesContext") = 2
        oFacePortExtractor1b.Parameter("LookingAxis") = 3
        oFacePortExtractor1b.Parameter("TrackFlag") = 2
        oFacePortExtractor1b.Parameter("Offset") = 0#
        oFacePortExtractor1b.Public = True
        oFacePortExtractor1b.Evaluate
        'Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort1"), oFacePortExtractor1b.ControlledInputs("Port"), sKeyOfMember1)
        
        Dim pPointAtSurfaceCenter1a As SP3DGeometricConstruction.GeometricConstruction
        Set pPointAtSurfaceCenter1a = pGCFactory.CreateEntity("PointAtSurfaceCenter", pPOM, "008d-PointAtSurfaceCenter")
        pPointAtSurfaceCenter1a.Inputs("Surface").Add oFacePortExtractor1a
        pPointAtSurfaceCenter1a.Evaluate
        
        Dim pPointAtSurfaceCenter1b As SP3DGeometricConstruction.GeometricConstruction
        Set pPointAtSurfaceCenter1b = pGCFactory.CreateEntity("PointAtSurfaceCenter", pPOM, "008d-PointAtSurfaceCenter")
        pPointAtSurfaceCenter1b.Inputs("Surface").Add oFacePortExtractor1b
        pPointAtSurfaceCenter1b.Evaluate
        
        Dim oParamDistMeasureParallel1 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel1 = pGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "008f-ParamDistMeasureParallel")
        oParamDistMeasureParallel1.Inputs("Graphics").Add pPointAtSurfaceCenter1a
        oParamDistMeasureParallel1.Inputs("Graphics").Add pPointAtSurfaceCenter1b
        oParamDistMeasureParallel1.Evaluate
        
        If CDbl(oParamDistMeasureParallel1.Parameter("Value")) > EPSILON Then
            Dim oSurfAtOffset1 As SP3DGeometricConstruction.GeometricConstruction
            Set oSurfAtOffset1 = pGCFactory.CreateEntity("SurfAtOffset", pPOM, "008g-SurfAtOffset")
            oSurfAtOffset1.Inputs("Surface").Add oFacePortExtractor1a
            oSurfAtOffset1.Inputs("TrackPoint").Add pPointAtSurfaceCenter1b
            oSurfAtOffset1.Parameter("Offset") = CDbl(oParamDistMeasureParallel1.Parameter("Value")) / 2
            oSurfAtOffset1.Evaluate
            
            Set oFacePortExtractor1a = oSurfAtOffset1
        End If
    End If
    
    Dim oCurveByIntersection1 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection1 = pGCFactory.CreateEntity("CurveByIntersection", pPOM, "009-CurveByIntersection")
    oCurveByIntersection1.Inputs("Surface1").Add oTrimmingPlane
    oCurveByIntersection1.Inputs("Surface2").Add oFacePortExtractor1a
    oCurveByIntersection1.Inputs("CoordinateSystem").Add pCoordinateSystem
    oCurveByIntersection1.Parameter("TrackFlag") = 1
    oCurveByIntersection1.Public = True
    oCurveByIntersection1.Evaluate

    Dim oPointAtCurveMiddle1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle1 = pGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "010-PointAtCurveMiddle")
    oPointAtCurveMiddle1.Inputs("Curve").Add oCurveByIntersection1
    oPointAtCurveMiddle1.Public = True
    oPointAtCurveMiddle1.Evaluate
    
    Dim oCurveByProjection2 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection2 = pGCFactory.CreateEntity("CurveByProjection", pPOM, "011-CurveByProjection")
    oCurveByProjection2.Inputs("Curve").Add oCurveByIntersection1
    oCurveByProjection2.Inputs("Surface").Add oFacePortExtractor2a
    oCurveByProjection2.Inputs("Line").Add oLineAxisPortExtractor1
    oCurveByProjection2.Inputs("CoordinateSystem").Add pCoordinateSystem
    oCurveByProjection2.Inputs("TrackPoint").Add oPointAtCurveMiddle1
    oCurveByProjection2.Parameter("TrackFlag") = 1
    oCurveByProjection2.Public = True
    On Error Resume Next
    Dim lError As Long
    oCurveByProjection2.Evaluate
    lError = Err.Number
    On Error GoTo 0
    If lError <> 0 Then Err.Raise CANNOT_BE_BOUNDED_BY_PLATE_SYSTEM
    

'    Dim lError As Long
'    On Error Resume Next: oCurveByProjection2.Evaluate: Let lError = Err.Number: On Error GoTo 0
'    If lError <> 0 Then
'        ' the face port is too small
'        If TypeOf oFacePortExtractor2a Is IJPlane Then
'            ' replace the face port by a big plane (100 m x 100 m)
'            Dim oCSByPlane7 As SP3DGeometricConstruction.GeometricConstruction
'            Set oCSByPlane7 = pGCFactory.CreateEntity("CSByPlane", pPOM, "011a-CSByPlane")
'            oCSByPlane7.Inputs("Plane").Add oFacePortExtractor2a
'            oCSByPlane7.Evaluate
'            oCSByPlane7.Public = True
'
'            Dim oPlaneFromCS7 As SP3DGeometricConstruction.GeometricConstruction
'            Set oPlaneFromCS7 = pGCFactory.CreateEntity("PlaneFromCS", pPOM, "011b-PlaneFromCS")
'            oPlaneFromCS7.Inputs("CoordinateSystem").Add oCSByPlane7
'            oPlaneFromCS7.Parameter("LookingAxis") = 3
'            oPlaneFromCS7.Parameter("Offset") = 0#
'            oPlaneFromCS7.Parameter("Range") = 100 ' -CDbl(oParamDistMeasureLength2.Parameter("Value")) ' 0.72
'            oPlaneFromCS7.Evaluate
'            oPlaneFromCS7.Public = True
'
'            Set oFacePortExtractor2a = oPlaneFromCS7
'        Else
'            Err.Raise lError
'        End If
'
'        oCurveByProjection2.Inputs("Surface").Add oFacePortExtractor2a
'        oCurveByProjection2.Evaluate
'    End If

    Dim oPointAtCurveMiddle2 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle2 = pGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "012-PointAtCurveMiddle")
    oPointAtCurveMiddle2.Inputs("Curve").Add oCurveByProjection2
    oPointAtCurveMiddle2.Public = True
    oPointAtCurveMiddle2.Evaluate

    Dim oExtRuledSurface5 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtRuledSurface5 = pGCFactory.CreateEntity("ExtRuledSurface", pPOM, "013-ExtRuledSurface")
    oExtRuledSurface5.Inputs("Curve1").Add oCurveByIntersection1
    oExtRuledSurface5.Inputs("Curve2").Add oCurveByProjection2
    oExtRuledSurface5.Parameter(sEXTENSION) = CDbl(pGeometricConstruction.Parameter(sEXTENSION))
    oExtRuledSurface5.Evaluate
    oExtRuledSurface5.Public = True

'    Dim oParamDistMeasureLength2 As SP3DGeometricConstruction.GeometricConstruction
'    Set oParamDistMeasureLength2 = pGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "014-ParamDistMeasureLength")
'    oParamDistMeasureLength2.Inputs("Curve").Add oCurveByProjection2
'    oParamDistMeasureLength2.Parameter("Value") = 0
'    oParamDistMeasureLength2.Evaluate
'    oParamDistMeasureLength2.Public = True
        
    ' orientate the support
    If True Then
        ' extract z axis of the coordinate system
        Dim oLineFromCS As GeometricConstruction
        Set oLineFromCS = pGCFactory.CreateEntity("LineFromCS", pPOM, "017-LineFromCS")
        oLineFromCS.Input("CoordinateSystem") = pCoordinateSystem
        oLineFromCS.Parameter("LookingAxis") = 3#
        oLineFromCS.Parameter("Length") = 1#
        oLineFromCS.Public = True
        oLineFromCS.Evaluate
        
        Call SurfaceBody_OrientateFollowingLine(oExtRuledSurface5.Output("Support", 1), oLineFromCS, CInt(pGeometricConstruction.Parameter(sLOCATION)) = 2)
    End If

    ' generate the outputs
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGeometricConstruction
    pGeometricConstructionMacro.Output("Support", 1) = oExtRuledSurface5.Output("Support", 1)
    pGeometricConstructionMacro.Output("Boundary", "TR1." + sKeyOfMember) = oTrimmingPlane
    If Not bIsContinuityPlate Then
        ' remove the previous boundary named "PlateSystem2"
'        Dim pObjectOfBoundary As IJDObject: Set pObjectOfBoundary = Nothing
'        On Error Resume Next: Set pObjectOfBoundary = pGeometricConstructionMacro.Outputs("Boundary")("PlateSystem2"): On Error GoTo 0
'        If Not pObjectOfBoundary Is Nothing Then
'            MsgBox "Remove boundary PlateSystem2"
'            pGeometricConstructionMacro.Outputs("Boundary").Remove ("PlateSystem2")
'            pObjectOfBoundary.Remove
'        End If
        pGeometricConstructionMacro.Output("Boundary", "FacePort2") = oFacePortExtractor2a.Output
    Else
'        ' remove the previous boundary named "FacePort2"
'        Set pObjectOfBoundary = Nothing
'        On Error Resume Next: Set pObjectOfBoundary = pGeometricConstructionMacro.Outputs("Boundary")("FacePort2"): On Error GoTo 0
'        If Not pObjectOfBoundary Is Nothing Then
'            MsgBox "Remove boundary FacePort2"
'            pGeometricConstructionMacro.Outputs("Boundary").Remove ("FacePort2")
'            pObjectOfBoundary.Remove
'        End If
        pGeometricConstructionMacro.Output("Boundary", "PlateSystem2") = oFacePortExtractor2a.Output
    End If
    pGeometricConstructionMacro.Output("Boundary", "FacePort3") = oFacePortExtractor3a.Output
    pGeometricConstructionMacro.Output("Boundary", "FacePort4") = oFacePortExtractor4a.Output
    Call GeometricConstructionMacro_CreateCoordinateSystemOfSupport(pGeometricConstructionMacro, pPOM, GetGCGeomFactory(), GetGCGeomFactory2(), pCoordinateSystem.Origin)
    Call GeometricConstructionMacro_CreateNode(pGeometricConstructionMacro, pPositionOfNode)
    
    Call DebugOut
End Sub
Private Function ModelBody_FromGeometry(pGeometry As IJDGeometry) As IJDModelBody
    ' initialize result
    Set ModelBody_FromGeometry = Nothing
    
    ' compute surface body
    Dim pModelBody As IJDModelBody
    If True Then
        If TypeOf pGeometry Is IJDModelBody Then
            Set pModelBody = pGeometry
        Else
            Dim pGeometryMisc As IJGeometryMisc
            Set pGeometryMisc = New DGeomOpsMisc
            Call pGeometryMisc.CreateModelGeometryFromGType(Nothing, pGeometry, Nothing, pModelBody)
        End If
    End If
    
    ' return result
    Set ModelBody_FromGeometry = pModelBody
End Function
'
' implementation of the IJGCMigrate interface
'
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
    Call DebugIn(Method("IJGCMigrate_Migrate"))

    ' retrieve position of node
    Dim pPositionOfNode As IJDPosition
    If True Then
        Dim MyGCMacro As IJGeometricConstructionMacro: Set MyGCMacro = MyGC
        Set pPositionOfNode = Position_FromPoint(MyGCMacro.Outputs("Node")(1))
    End If
    
    ' loop on all the inputs, whose name is prefixed by sMEMBER_PART
    Dim pGeometricConstructionDefinition As IJGeometricConstructionDefinition: Set pGeometricConstructionDefinition = MyGC.definition
    Dim i  As Integer
    For i = 1 To pGeometricConstructionDefinition.InputCount
        Dim sName As String, sPrompt As String, sFilter As String, lMinConnected As Long, lMaxConnected As Long, sComputeIIDs As String
        Call pGeometricConstructionDefinition.GetInputInfoByIndex(i, sName, sPrompt, sFilter, lMinConnected, lMaxConnected, sComputeIIDs)
                
        ' migrate these inputs, if one of them has been replaced
        If Mid(sName, 1, Len(sMEMBER_PART)) = sMEMBER_PART _
        Or sName = sREFERENCE_MEMBER_PART Then Call GeometricConstruction_MigrateInputs(MyGC, pMigrateHelper, sName, pPositionOfNode)
    Next
    
    Call DebugOut
End Sub
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    ' Delegate to the APS when exists
    Set IJGCToDoDelegate_ToDoDelegate = Nothing
    If pGC.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        Set IJGCToDoDelegate_ToDoDelegate = pGC.ControlledInputs("AdvancedPlateSystem")(1)
    End If
End Property

