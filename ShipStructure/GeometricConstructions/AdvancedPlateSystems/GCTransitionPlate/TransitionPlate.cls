VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransitionPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum Errors
    MISSING_MEMBER_PART1 = 1
    MISSING_MEMBER_PART2 = 2
    MISSING_SIDE_PLATE1 = 3
    MISSING_SIDE_PLATE2 = 4
    DUPLICATE_INPUTS = 5
    CANNOT_RETRIEVE_TRIMMING_PLANE1 = 6
    CANNOT_RETRIEVE_TRIMMING_PLANE2 = 7
End Enum

Private Const sMEMBER_PART1  As String = "MemberPart1"
Private Const sMEMBER_PART2  As String = "MemberPart2"
Private Const sSIDE_PLATE1  As String = "SidePlate1"
Private Const sSIDE_PLATE2  As String = "SidePlate2"
Private Const sEXTENSION  As String = "Extension"
Private Const sSIDE  As String = "Side"

Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGCMigrate
Implements IJTransitionPlate
Implements IJGCToDoDelegate
Implements IJGCConnectedElementSC

Sub BreakPoint()
    Dim i As Integer: Let i = 0
End Sub
Private Property Get Source() As String
    Let Source = "GCTransitionPlate.TransitionPlate"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property
Private Sub Class_Initialize()
    Call DebugStart(Source)

'    Let m_lCountOfDebugSupport = m_lCountOfDebugSupport + 1
'    Set m_pDebugSupport(m_lCountOfDebugSupport) = New DebugSupport
'    Let m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_SOURCE = "TransitionPlate1 (" + CStr(m_lCountOfDebugSupport) + ")"
'    Call m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_MSG(">> TransitionPlate1")
End Sub
Private Sub Class_Terminate()
'    Call m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_MSG("<< TransitionPlate1")
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

        Dim bAPSTransformed As Boolean
        Dim bInputMemberTransformed1 As Boolean
        Dim bInputMemberTransformed2 As Boolean
        bAPSTransformed = pInfo.IsDesignObjectInTransformSet(oPlateSystemAdvanced)
        If bAPSTransformed Then
            'APS transformed, unbound members that are not transformed
            Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGC
            
            ' retrieve info about the input member #1
            Dim pElementsOfMembers1 As IJElements: Set pElementsOfMembers1 = Elements_GetValidElements(pGC.Inputs(sMEMBER_PART1))
            Dim iCountOfMembers1 As Integer:  Let iCountOfMembers1 = pElementsOfMembers1.Count
            Dim sKeysOfMembers1() As String: If iCountOfMembers1 > 0 Then Let sKeysOfMembers1 = Elements_GetDummyKeys(pElementsOfMembers1)
            
             ' retrieve info about the input member #2
            Dim pElementsOfMembers2 As IJElements: Set pElementsOfMembers2 = Elements_GetValidElements(pGC.Inputs(sMEMBER_PART2))
            Dim iCountOfMembers2 As Integer:  Let iCountOfMembers2 = pElementsOfMembers2.Count
            Dim sKeysOfMembers2() As String: If iCountOfMembers2 > 0 Then Let sKeysOfMembers2 = Elements_GetDummyKeys(pElementsOfMembers2)
         
            If iCountOfMembers1 = 1 Then
                bInputMemberTransformed1 = pInfo.IsDesignObjectInTransformSet(pElementsOfMembers1(1))
            End If
            If iCountOfMembers2 = 1 Then
                bInputMemberTransformed2 = pInfo.IsDesignObjectInTransformSet(pElementsOfMembers2(1))
            End If
           
            If Not bInputMemberTransformed1 Or Not bInputMemberTransformed2 Then
                ' retrieve info about SidePlates
                Dim oSidePlate1 As Object
                Dim oSidePlate2 As Object
                If True Then
                    Dim pElementsOfValidElements As IJElements:
                    Set pElementsOfValidElements = Elements_GetValidElements(pGC.Inputs(sSIDE_PLATE1))
                    If pElementsOfValidElements.Count = 1 Then Set oSidePlate1 = pElementsOfValidElements(1)
                    
                    Set pElementsOfValidElements = Elements_GetValidElements(pGC.Inputs(sSIDE_PLATE2))
                    If pElementsOfValidElements.Count = 1 Then Set oSidePlate2 = pElementsOfValidElements(1)
                End If
                
                 ' retrieve info about boundaries
                Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
                Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
                Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)
    
                ' compute the node
                Dim pPositionOfNode As IJDPosition
                If iCountOfMembers1 = 1 _
                And iCountOfMembers2 = 1 _
                And Not oSidePlate1 Is Nothing _
                And Not oSidePlate2 Is Nothing Then
                    ' if something goes wrong, we will at least try to retrieve the node
                    On Error Resume Next
                    
                     ' get POM for debug
                    Dim pPOM As IJDPOM: ' Set pPOM = GetPOMFromObject(pGeometricConstruction)
                    ' get factory
                    Dim pGCFactory As IJGeometricConstructionEntitiesFactory
                    Set pGCFactory = New GeometricConstructionEntitiesFactory
                   
                    ' compute the coordinate system
                    Dim pCoordinateSystem As IJDCoordinateSystem
                    Set pCoordinateSystem = ComputeCS(pGC, pPOM, pGCFactory, _
                                                      pElementsOfMembers1(1), sKeysOfMembers1(1), _
                                                      pElementsOfMembers2(1), sKeysOfMembers2(1), _
                                                      oSidePlate1, _
                                                      oSidePlate2)
                    If Not pCoordinateSystem Is Nothing Then Set pPositionOfNode = pCoordinateSystem.Origin
                    
                    On Error GoTo 0
                End If
                
                ' retrieve the node, if not computed
                If pPositionOfNode Is Nothing Then Set pPositionOfNode = GeometricConstructionMacro_RetrievePositionOfNode(pGeometricConstructionMacro)
                
                ' un-bound removed members
                If Not bInputMemberTransformed1 Then
                     Call Member_UnBoundFromEdgeOfAPS(pElementsOfMembers1(1), pGC, oPlateSystemAdvanced, pPositionOfNode)
                End If
                If Not bInputMemberTransformed2 Then
                     Call Member_UnBoundFromEdgeOfAPS(pElementsOfMembers2(1), pGC, oPlateSystemAdvanced, pPositionOfNode)
                End If
            End If
        End If
    End If
End Sub

'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Initialize"))
    
    ' located input
    Call pGeometricConstructionDefinition.AddInput(sMEMBER_PART1, "Select first Member", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput(sMEMBER_PART2, "Select opposite Member", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    ' limit next 2 inputs to be any type of APS
    Call pGeometricConstructionDefinition.AddInput(sSIDE_PLATE1, "Select the back SidePlate (APS)", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput(sSIDE_PLATE2, "Select the front SidePlate (APS)", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    ' controlled inputs
    Call pGeometricConstructionDefinition.AddControlledInput("AdvancedPlateSystem", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort2")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort2")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort3")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort4")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort3a")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort4a")

    ' special controlled input to indicate that some PlateSystems of incoming Members have not been correctly un-bounded during the disconnect semantic,
    ' in the case, where the APS and some incoming Members are simultaneously deleted.
    ' this is due to the fact that FrameConnections depending from a deleted Member are first deleted, before being later recreated.
    ' our disconnect semantic executes, after those FrameConnections get deleted and before they get re-created.
    ' we have to wait the triggering of our compute semantic to have access again to the re-created ones.
    Call pGeometricConstructionDefinition.AddControlledInput("UnBoundedPlateSystems", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    ' parameters
    Call pGeometricConstructionDefinition.AddParameter("Extension", "Extension", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    
    ' code listed parameters
    Call pGeometricConstructionDefinition.AddParameter("Side", "Side", GCCodeList, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Side", "Bottom", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Side", "Top", 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Side", "Centered", 3)
        
    ' error codes
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MEMBER_PART1, "MissingMemberPart1", "Missing the member part #1")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MEMBER_PART2, "MissingMemberPart2", "Missing the member part #2")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_SIDE_PLATE1, "MissingSidePlate1", "Missing the side plate (APS) #1")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_SIDE_PLATE2, "MissingSidePlate2", "Missing the side plate (APS) #2")
    Call pGeometricConstructionDefinition.AddErrorValue(DUPLICATE_INPUTS, "DuplicateInputs", "2 inputs are the same")
    Call pGeometricConstructionDefinition.AddErrorValue(CANNOT_RETRIEVE_TRIMMING_PLANE1, "CannotRetrieveTrimmingPlane1", "Cannot retrieve trimming plane along MemberPart1 from APS SidePlate #1")
    Call pGeometricConstructionDefinition.AddErrorValue(CANNOT_RETRIEVE_TRIMMING_PLANE2, "CannotRetrieveTrimmingPlane2", "Cannot retrieve trimming plane along MemberPart2 from APS SidePlate #2")
    
    ' outputs
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Support")
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypePoint3d, "Node")
    Call pGeometricConstructionDefinition.AddOutput(GCLocalCoordinateSystem, "CoordinateSystem")
    
    Call DebugOut
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler: Call BreakPoint
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Evaluate"))

    ' use GC as GCMacro
    Dim pGeometricConstructonMacro As IJGeometricConstructionMacro: Set pGeometricConstructonMacro = pGeometricConstruction
    
    ' continue the work partially done during the disconnect semantic
    If pGeometricConstruction.ControlledInputs("UnBoundedPlateSystems").Count > 0 Then
        On Error Resume Next
        Call Members_UnboundOnDelete(pGeometricConstruction, pPOM, Nothing)
        On Error GoTo 0
        
        ' delete myself, because the APS is deleted and we do not want any more the compute semantic to fire
        Call Object_Delete(pGeometricConstruction)
        Exit Sub
    End If
    
    ' get factory
    Dim pGCFactory As IJGeometricConstructionEntitiesFactory
    Set pGCFactory = New GeometricConstructionEntitiesFactory
    
    ' check for errors on inputs
    If True Then
        ' check mandatory inputs
        If pGeometricConstruction.Inputs(sMEMBER_PART1).Count = 0 Then Err.Raise MISSING_MEMBER_PART1
        If pGeometricConstruction.Inputs(sMEMBER_PART2).Count = 0 Then Err.Raise MISSING_MEMBER_PART2
        If pGeometricConstruction.Inputs(sSIDE_PLATE1).Count = 0 Then Err.Raise MISSING_SIDE_PLATE1
        If pGeometricConstruction.Inputs(sSIDE_PLATE2).Count = 0 Then Err.Raise MISSING_SIDE_PLATE2
            
        ' check duplicate inputs
        Dim pElements As IJElements: Set pElements = New JObjectCollection
        Call pElements.Add(pGeometricConstruction.Input(sMEMBER_PART1))
        If pElements.Contains(pGeometricConstruction.Input(sMEMBER_PART2)) Then Err.Raise DUPLICATE_INPUTS
        Call pElements.Add(pGeometricConstruction.Input(sMEMBER_PART2))
        If pElements.Contains(pGeometricConstruction.Input(sSIDE_PLATE1)) Then Err.Raise DUPLICATE_INPUTS
        Call pElements.Add(pGeometricConstruction.Input(sSIDE_PLATE1))
        If pElements.Contains(pGeometricConstruction.Input(sSIDE_PLATE2)) Then Err.Raise DUPLICATE_INPUTS
    End If
    
    ' retrieve info about the input member #1
    Dim pElementsOfMembers1 As IJElements: Set pElementsOfMembers1 = pGeometricConstruction.Inputs(sMEMBER_PART1)
    Dim iCountOfMembers1 As Integer:  Let iCountOfMembers1 = pElementsOfMembers1.Count
    Dim sKeysOfMembers1() As String: If iCountOfMembers1 > 0 Then Let sKeysOfMembers1 = Elements_GetKeys(pElementsOfMembers1)
    
    ' retrieve info about the input member #2
    Dim pElementsOfMembers2 As IJElements: Set pElementsOfMembers2 = pGeometricConstruction.Inputs(sMEMBER_PART2)
    Dim iCountOfMembers2 As Integer:  Let iCountOfMembers2 = pElementsOfMembers2.Count
    Dim sKeysOfMembers2() As String: If iCountOfMembers2 > 0 Then Let sKeysOfMembers2 = Elements_GetKeys(pElementsOfMembers2)
        
    ' retrieve info about boundaries
    Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructonMacro.Outputs("Boundary")
    Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
    Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)
    
    ' retrieve SidePlates
    Dim oSidePlate1 As Object: Set oSidePlate1 = pGeometricConstruction.Input(sSIDE_PLATE1)
    Dim oSidePlate2 As Object: Set oSidePlate2 = pGeometricConstruction.Input(sSIDE_PLATE2)
    
    ' retrieve TrimmingPlanes
    Dim oTrimmingPlane1 As Object
    Dim oTrimmingPlane2 As Object
    Call GetTrimmingPlanes(pPOM, pElementsOfMembers1(1), pElementsOfMembers2(1), oSidePlate1, oSidePlate2, oTrimmingPlane1, oTrimmingPlane2)
    
    ' compute the coordinate system
    Dim pCoordinateSystem As IJDCoordinateSystem: Set pCoordinateSystem = ComputeCS(pGeometricConstruction, pPOM, pGCFactory, _
                                                                                    pElementsOfMembers1(1), sKeysOfMembers1(1), _
                                                                                    pElementsOfMembers2(1), sKeysOfMembers2(1), _
                                                                                    oSidePlate1, oSidePlate2)

    ' un-bound removed members
    If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        Call Members_UnboundFromEdgeOfAPS(pElementsOfMembers1, iCountOfMembers1, sKeysOfMembers1, _
                                          pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                          pGeometricConstruction, 1, pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1), pCoordinateSystem.Origin)
                                         
        Call Members_UnboundFromEdgeOfAPS(pElementsOfMembers2, iCountOfMembers2, sKeysOfMembers2, _
                                          pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                          pGeometricConstruction, 2, pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1), pCoordinateSystem.Origin)
    End If
        
    ' compute outputs
    Call ComputeOutputs(pGeometricConstruction, pPOM, pGCFactory, _
                        pCoordinateSystem, _
                        pElementsOfMembers1(1), sKeysOfMembers1(1), _
                        pElementsOfMembers2(1), sKeysOfMembers2(1), _
                        oSidePlate1, oSidePlate2, _
                        oTrimmingPlane1, oTrimmingPlane2)
    
    If iCountOfBoundaries > 0 Then
        Call Keys_RemoveKey(sKeysOfBoundaries, "TR1." + sKeysOfMembers1(1))
        Call Keys_RemoveKey(sKeysOfBoundaries, "TR2." + sKeysOfMembers2(1))
        Call Keys_RemoveKey(sKeysOfBoundaries, "FacePort3")
        Call Keys_RemoveKey(sKeysOfBoundaries, "FacePort4")
    End If
        
    ' trim members
    If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        ' get the key of the APS
        Dim sKey As String: Let sKey = pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").GetKey(pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1))
        
        ' if the key is not set to "NoBound" then do the bound (this means that the APS has already computed its geometry and edge ports can be extracted)
        If Not sKey = "NotBounded" And Not sKey = "JustMigrated" Then
            Call Members_BoundByOrUnboundFromEdgeOfAPS(pElementsOfMembers1, iCountOfMembers1, sKeysOfMembers1, _
                                                       pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                                       pGeometricConstruction, 1, pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1), pCoordinateSystem.Origin)
            
            Call Members_BoundByOrUnboundFromEdgeOfAPS(pElementsOfMembers2, iCountOfMembers2, sKeysOfMembers2, _
                                                       pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                                       pGeometricConstruction, 2, pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1), pCoordinateSystem.Origin)
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
' implementation of the IJTransitionPlate interface
'
Private Function IJTransitionPlate_Create(ByVal pPOM As IJDPOM, _
                                          ByVal oMemberPart1 As Object, _
                                          ByVal oMemberPart2 As Object, _
                                          ByVal oSidePlate1 As Object, _
                                          ByVal oSidePlate2 As Object, _
                                          Optional ByVal dExtension As Double = 0#, _
                                          Optional ByVal eTransitionPlateLocation As eTransitionPlateLocation = 1) As Object
    ' create factory
    Dim pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory
    Set pGeometricConstructionFactory = New GeometricConstructionEntitiesFactory

    ' create geometric construction
    Dim pGeometricConstruction As IJGeometricConstruction
    Set pGeometricConstruction = pGeometricConstructionFactory.CreateEntity("TransitionPlate", pPOM)
    
    ' provide inputs and parameters
    Call pGeometricConstruction.Inputs(sMEMBER_PART1).Add(oMemberPart1, "1")
    Call pGeometricConstruction.Inputs(sMEMBER_PART2).Add(oMemberPart2, "1")
    Call pGeometricConstruction.Inputs(sSIDE_PLATE1).Add(oSidePlate1, "1")
    Call pGeometricConstruction.Inputs(sSIDE_PLATE2).Add(oSidePlate2, "1")
    pGeometricConstruction.Parameter(sEXTENSION) = dExtension
    pGeometricConstruction.Parameter(sSIDE) = eTransitionPlateLocation
    
    ' return geometric construction
    Set IJTransitionPlate_Create = pGeometricConstruction
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
    ElseIf Mid(pRelationship.Name, 1, Len(sMEMBER_PART1)) = sMEMBER_PART1 And bIsDestinationDeleted Then
        ' an input member is deleting, disconnect the corresponding ports
        Dim sKey As String: Let sKey = Mid(pRelationship.Name, Len(sMEMBER_PART1) + 2)
        pGeometricConstruction.ControlledInputs("AxisPort1").Remove (sKey)
        pGeometricConstruction.ControlledInputs("FacePort1").Remove (sKey)
    ElseIf Mid(pRelationship.Name, 1, Len(sMEMBER_PART2)) = sMEMBER_PART2 And bIsDestinationDeleted Then
        ' an input member is deleting, disconnect the corresponding ports
        Let sKey = Mid(pRelationship.Name, Len(sMEMBER_PART2) + 2)
        pGeometricConstruction.ControlledInputs("AxisPort2").Remove (sKey)
        pGeometricConstruction.ControlledInputs("FacePort2").Remove (sKey)
    End If
    Exit Sub
    
    Call DebugOut
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
        
    ' retrieve info about the input member #1
    Dim pElementsOfMembers1 As IJElements: Set pElementsOfMembers1 = Elements_GetValidElements(pGeometricConstruction.Inputs(sMEMBER_PART1))
    Dim iCountOfMembers1 As Integer:  Let iCountOfMembers1 = pElementsOfMembers1.Count
    Dim sKeysOfMembers1() As String: If iCountOfMembers1 > 0 Then Let sKeysOfMembers1 = Elements_GetDummyKeys(pElementsOfMembers1)
    
    ' retrieve info about the input member #2
    Dim pElementsOfMembers2 As IJElements: Set pElementsOfMembers2 = Elements_GetValidElements(pGeometricConstruction.Inputs(sMEMBER_PART2))
    Dim iCountOfMembers2 As Integer:  Let iCountOfMembers2 = pElementsOfMembers2.Count
    Dim sKeysOfMembers2() As String: If iCountOfMembers2 > 0 Then Let sKeysOfMembers2 = Elements_GetDummyKeys(pElementsOfMembers2)
    
    ' retrieve info about SidePlates
    Dim oSidePlate1 As Object
    Dim oSidePlate2 As Object
    If True Then
        Dim pElementsOfValidElements As IJElements:
        Set pElementsOfValidElements = Elements_GetValidElements(pGeometricConstruction.Inputs(sSIDE_PLATE1))
        If pElementsOfValidElements.Count = 1 Then Set oSidePlate1 = pElementsOfValidElements(1)
        
        Set pElementsOfValidElements = Elements_GetValidElements(pGeometricConstruction.Inputs(sSIDE_PLATE2))
        If pElementsOfValidElements.Count = 1 Then Set oSidePlate2 = pElementsOfValidElements(1)
    End If
    
    If iCountOfMembers1 = 1 _
    Or iCountOfMembers2 = 1 Then
        ' retrieve info about boundaries
        Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
        Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
        Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)
    
        ' compute the node
        Dim pPositionOfNode As IJDPosition
        If iCountOfMembers1 = 1 _
        And iCountOfMembers2 = 1 _
        And Not oSidePlate1 Is Nothing _
        And Not oSidePlate2 Is Nothing Then
            ' if something goes wrong, we will at least try to retrieve the node
            On Error Resume Next
            
            ' compute the coordinate system
            Dim pCoordinateSystem As IJDCoordinateSystem
            Set pCoordinateSystem = ComputeCS(pGeometricConstruction, pPOM, pGCFactory, _
                                              pElementsOfMembers1(1), sKeysOfMembers1(1), _
                                              pElementsOfMembers2(1), sKeysOfMembers2(1), _
                                              oSidePlate1, _
                                              oSidePlate2)
            If Not pCoordinateSystem Is Nothing Then Set pPositionOfNode = pCoordinateSystem.Origin
            
            On Error GoTo 0
        End If
        
        ' retrieve the node, if not computed
        If pPositionOfNode Is Nothing Then Set pPositionOfNode = GeometricConstructionMacro_RetrievePositionOfNode(pGeometricConstructionMacro)
        
        ' un-bound removed members
        If iCountOfMembers1 = 1 Then _
            Call Members_UnboundFromEdgeOfAPS(pElementsOfMembers1, iCountOfMembers1, sKeysOfMembers1, _
                                              pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                              pGeometricConstruction, 1, oAdvancedPlateSystem, pPositionOfNode)
                                         
        If iCountOfMembers2 = 1 Then _
            Call Members_UnboundFromEdgeOfAPS(pElementsOfMembers2, iCountOfMembers2, sKeysOfMembers2, _
                                              pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                              pGeometricConstruction, 2, oAdvancedPlateSystem, pPositionOfNode)
    End If
    
    Call DebugOut
End Sub
Private Function ComputeCS(pGeometricConstruction As IJGeometricConstruction, pPOM As IJDPOM, pGCFactory As IJGeometricConstructionEntitiesFactory, _
                           oMember1, sKeyOfMember1 As String, _
                           oMember2, sKeyOfMember2 As String, _
                           oSidePlate1 As Object, _
                           oSidePlate2 As Object) As IJDCoordinateSystem
    Call DebugIn(Method("ComputeCS"))
    
    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = pGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add oMember1
    oLineAxisPortExtractor1.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("AxisPort1")
    oLineAxisPortExtractor1.Public = True
    oLineAxisPortExtractor1.Evaluate
    Call Elements_ReplaceElementsWithKey(pGeometricConstruction.ControlledInputs("AxisPort1"), oLineAxisPortExtractor1.ControlledInputs("Port"), sKeyOfMember1)

    Dim oLineAxisPortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor2 = pGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "002-LineAxisPortExtractor")
    oLineAxisPortExtractor2.Inputs("MemberPart").Add oMember2
    oLineAxisPortExtractor2.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("AxisPort2")
    oLineAxisPortExtractor2.Public = True
    oLineAxisPortExtractor2.Evaluate
    Call Elements_ReplaceElementsWithKey(pGeometricConstruction.ControlledInputs("AxisPort2"), oLineAxisPortExtractor2.ControlledInputs("Port"), sKeyOfMember2)

    Dim oPointAtCurveMiddle1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle1 = pGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "003-PointAtCurveMiddle")
    oPointAtCurveMiddle1.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveMiddle1.Public = True
    oPointAtCurveMiddle1.Evaluate

    Dim oPointAtCurveMiddle2 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle2 = pGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "004-PointAtCurveMiddle")
    oPointAtCurveMiddle2.Inputs("Curve").Add oLineAxisPortExtractor2
    oPointAtCurveMiddle2.Public = True
    oPointAtCurveMiddle2.Evaluate

    Dim oPointAtCurveExtremity1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity1 = pGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "005-PointAtCurveExtremity")
    oPointAtCurveExtremity1.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveExtremity1.Inputs("TrackPoint").Add oPointAtCurveMiddle2
    oPointAtCurveExtremity1.Parameter("TrackFlag") = 1
    oPointAtCurveExtremity1.Public = True
    oPointAtCurveExtremity1.Evaluate

    Dim oFacePortExtractor3 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor3 = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "006-FacePortExtractor")
    oFacePortExtractor3.Inputs("Connectable").Add oSidePlate1
    oFacePortExtractor3.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort3")
    oFacePortExtractor3.Parameter("GeometrySelector") = 4
    oFacePortExtractor3.Parameter("FacesContext") = 1
    oFacePortExtractor3.Parameter("LookingAxis") = 2
    oFacePortExtractor3.Parameter("TrackFlag") = 1
    oFacePortExtractor3.Parameter("Offset") = 0
    oFacePortExtractor3.Public = True
    oFacePortExtractor3.Evaluate
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort3"), oFacePortExtractor3.ControlledInputs("Port"))
   
    Dim oFacePortExtractor4 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor4 = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "007-FacePortExtractor")
    oFacePortExtractor4.Inputs("Connectable").Add oSidePlate2
    oFacePortExtractor4.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort4")
    oFacePortExtractor4.Parameter("GeometrySelector") = 4
    oFacePortExtractor4.Parameter("FacesContext") = 1
    oFacePortExtractor4.Parameter("LookingAxis") = 2
    oFacePortExtractor4.Parameter("TrackFlag") = 1
    oFacePortExtractor4.Parameter("Offset") = 0
    oFacePortExtractor4.Public = True
    oFacePortExtractor4.Evaluate
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort4"), oFacePortExtractor4.ControlledInputs("Port"))
   
       ' project P1 on F3
    Dim oCSByPlane3 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane3 = pGCFactory.CreateEntity("CSByPlane", pPOM, "008-CSByPlane")
    oCSByPlane3.Inputs("Plane").Add oFacePortExtractor3
    oCSByPlane3.Inputs("Point").Add oPointAtCurveExtremity1
    oCSByPlane3.Public = True
    oCSByPlane3.Evaluate

    Dim oLineFromCS3x As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS3x = pGCFactory.CreateEntity("LineFromCS", pPOM, "009-LineFromCS")
    oLineFromCS3x.Inputs("CoordinateSystem").Add oCSByPlane3
    oLineFromCS3x.Parameter("LookingAxis") = 1#
    oLineFromCS3x.Parameter("Length") = -1#
    oLineFromCS3x.Public = True
    oLineFromCS3x.Evaluate

    Dim oLineFromCS3z As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS3z = pGCFactory.CreateEntity("LineFromCS", pPOM, "010-LineFromCS")
    oLineFromCS3z.Inputs("CoordinateSystem").Add oCSByPlane3
    oLineFromCS3z.Parameter("LookingAxis") = 3#
    oLineFromCS3z.Parameter("Length") = 1#
    oLineFromCS3z.Public = True
    oLineFromCS3z.Evaluate

    Dim oCurveByProjection4a As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection4a = pGCFactory.CreateEntity("CurveByProjection", pPOM, "011-CurveByProjection")
    oCurveByProjection4a.Inputs("Curve").Add oLineFromCS3x
    oCurveByProjection4a.Inputs("Surface").Add oFacePortExtractor3
    oCurveByProjection4a.Inputs("Line").Add oLineFromCS3z
    oCurveByProjection4a.Parameter("TrackFlag") = 1
    oCurveByProjection4a.Public = True
    oCurveByProjection4a.Evaluate

    Dim oCurveByProjection4b As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection4b = pGCFactory.CreateEntity("CurveByProjection", pPOM, "012-CurveByProjection")
    oCurveByProjection4b.Inputs("Curve").Add oLineFromCS3x
    oCurveByProjection4b.Inputs("Surface").Add oFacePortExtractor4
    oCurveByProjection4b.Inputs("Line").Add oLineFromCS3z
    oCurveByProjection4b.Parameter("TrackFlag") = 1
    oCurveByProjection4b.Public = True
    oCurveByProjection4b.Evaluate

    Dim oPointAtCurveMiddle4a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle4a = pGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "013-PointAtCurveMiddle")
    oPointAtCurveMiddle4a.Inputs("Curve").Add oCurveByProjection4a
    oPointAtCurveMiddle4a.Public = True
    oPointAtCurveMiddle4a.Evaluate

    Dim oPointAtCurveMiddle4b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle4b = pGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "014-PointAtCurveMiddle")
    oPointAtCurveMiddle4b.Inputs("Curve").Add oCurveByProjection4b
    oPointAtCurveMiddle4b.Public = True
    oPointAtCurveMiddle4b.Evaluate

    Dim oLineByPoints5z As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints5z = pGCFactory.CreateEntity("LineByPoints", pPOM, "015-LineByPoints")
    oLineByPoints5z.Inputs("StartPoint").Add oPointAtCurveMiddle4a
    oLineByPoints5z.Inputs("EndPoint").Add oPointAtCurveMiddle4b
    oLineByPoints5z.Evaluate
    oLineByPoints5z.Public = True
    
    Dim oParamDistMeasureLength5z As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength5z = pGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "016-ParamDistMeasureLength")
    oParamDistMeasureLength5z.Inputs("Curve").Add oLineByPoints5z
    oParamDistMeasureLength5z.Evaluate
    oParamDistMeasureLength5z.Public = True
    
    Dim oPointAlongCurve5z As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve5z = pGCFactory.CreateEntity("PointAlongCurve", pPOM, "017-PointAlongCurve")
    oPointAlongCurve5z.Inputs("Curve").Add oLineByPoints5z
    oPointAlongCurve5z.Inputs("Point").Add oPointAtCurveMiddle4b
    oPointAlongCurve5z.Parameter("Distance") = 2 * CDbl(oParamDistMeasureLength5z.Parameter("Value"))
    oPointAlongCurve5z.Evaluate
    oPointAlongCurve5z.Public = True
    
    Dim oLineByPoints5zz As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints5zz = pGCFactory.CreateEntity("LineByPoints", pPOM, "018-LineByPoints")
    oLineByPoints5zz.Inputs("StartPoint").Add oPointAtCurveMiddle4a
    oLineByPoints5zz.Inputs("EndPoint").Add oPointAlongCurve5z
    oLineByPoints5zz.Evaluate
    oLineByPoints5zz.Public = True
    
    Dim oLineByPoints5x As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints5x = pGCFactory.CreateEntity("LineByPoints", pPOM, "019-LineByPoints")
    oLineByPoints5x.Inputs("StartPoint").Add oPointAtCurveExtremity1
    oLineByPoints5x.Inputs("EndPoint").Add oPointAtCurveMiddle2
    oLineByPoints5x.Evaluate
    oLineByPoints5x.Public = True
    
'''    Dim oCSByPoints3 As SP3DGeometricConstruction.GeometricConstruction
'''    Set oCSByPoints3 = pGCFactory.CreateEntity("CSByPoints", pPOM, "013-CSByPoints")
'''    oCSByPoints3.Inputs("OriginPoint").Add oPointAtCurveExtremity1
'''    oCSByPoints3.Inputs("AxisPoint1").Add oPointAtCurveMiddle4
'''    oCSByPoints3.Inputs("AxisPoint2").Add oPointAtCurveMiddle2
'''    oCSByPoints3.Parameter("AxesRoles") = 3
'''    oCSByPoints3.Parameter("CSOrientation") = 1
'''    oCSByPoints3.Public = True
'''    oCSByPoints3.Evaluate
    
    Dim oCSByLines3 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines3 = pGCFactory.CreateEntity("CSByLines", pPOM, "020-CSByLines")
    oCSByLines3.Inputs("AxisLine1").Add oLineByPoints5zz
    oCSByLines3.Inputs("AxisLine2").Add oLineByPoints5x
    oCSByLines3.Parameter("AxesRoles") = 3
    oCSByLines3.Parameter("CSOrientation") = 1
    oCSByLines3.Public = True
    oCSByLines3.Evaluate
    
    Set ComputeCS = oCSByLines3
    
    Call DebugOut
End Function
Private Sub ComputeOutputs(pGeometricConstruction As IJGeometricConstruction, pPOM As IJDPOM, pGCFactory As IJGeometricConstructionEntitiesFactory, _
                           pCoordinateSystem As IJDCoordinateSystem, _
                           oMember1 As Object, sKeyOfMember1 As String, _
                           oMember2 As Object, sKeyOfMember2 As String, _
                           oSidePlate1 As Object, _
                           oSidePlate2 As Object, _
                           oTrimmingPlane1 As Object, _
                           oTrimmingPlane2 As Object)
    Call DebugIn(Method("ComputeOutputs"))
    
    Dim oFacePortExtractor3a As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor3a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "021-FacePortExtractor")
    oFacePortExtractor3a.Inputs("Connectable").Add oSidePlate1
    oFacePortExtractor3a.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort3a")
    oFacePortExtractor3a.Parameter("GeometrySelector") = 2
    oFacePortExtractor3a.Parameter("FacesContext") = 1
    oFacePortExtractor3a.Parameter("LookingAxis") = 2
    oFacePortExtractor3a.Parameter("TrackFlag") = 1
    oFacePortExtractor3a.Parameter("Offset") = 0
    oFacePortExtractor3a.Public = True
    oFacePortExtractor3a.Evaluate
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort3a"), oFacePortExtractor3a.ControlledInputs("Port"))
  
    Dim oFacePortExtractor4a As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor4a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "022-FacePortExtractor")
    oFacePortExtractor4a.Inputs("Connectable").Add oSidePlate2
    oFacePortExtractor4a.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort4a")
    oFacePortExtractor4a.Parameter("GeometrySelector") = 2
    oFacePortExtractor4a.Parameter("FacesContext") = 1
    oFacePortExtractor4a.Parameter("LookingAxis") = 2
    oFacePortExtractor4a.Parameter("TrackFlag") = 1
    oFacePortExtractor4a.Parameter("Offset") = 0
    oFacePortExtractor4a.Public = True
    oFacePortExtractor4a.Evaluate
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort4a"), oFacePortExtractor4a.ControlledInputs("Port"))

    If CInt(pGeometricConstruction.Parameter("Side")) <> 3 Then
        Dim oFacePortExtractor1a As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor1a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "023a-FacePortExtractor")
        oFacePortExtractor1a.Inputs("Connectable").Add oMember1
        oFacePortExtractor1a.Inputs("CoordinateSystem").Add pCoordinateSystem
        oFacePortExtractor1a.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort1")
        oFacePortExtractor1a.Parameter("GeometrySelector") = 4
        oFacePortExtractor1a.Parameter("FacesContext") = 2
        oFacePortExtractor1a.Parameter("LookingAxis") = 2
        oFacePortExtractor1a.Parameter("TrackFlag") = CInt(pGeometricConstruction.Parameter("Side"))
        oFacePortExtractor1a.Parameter("Offset") = 0#
        oFacePortExtractor1a.Public = True
        oFacePortExtractor1a.Evaluate
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort1"), oFacePortExtractor1a.ControlledInputs("Port"), sKeyOfMember1)
    Else
        Set oFacePortExtractor1a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "023a-FacePortExtractor")
        oFacePortExtractor1a.Inputs("Connectable").Add oMember1
        oFacePortExtractor1a.Inputs("CoordinateSystem").Add pCoordinateSystem
        oFacePortExtractor1a.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort1")
        oFacePortExtractor1a.Parameter("GeometrySelector") = 4
        oFacePortExtractor1a.Parameter("FacesContext") = 2
        oFacePortExtractor1a.Parameter("LookingAxis") = 2
        oFacePortExtractor1a.Parameter("TrackFlag") = 1
        oFacePortExtractor1a.Parameter("Offset") = 0#
        oFacePortExtractor1a.Public = True
        oFacePortExtractor1a.Evaluate
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort1"), oFacePortExtractor1a.ControlledInputs("Port"), sKeyOfMember1)
        
        Dim oFacePortExtractor1b As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor1b = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "023b-FacePortExtractor")
        oFacePortExtractor1b.Inputs("Connectable").Add oMember1
        oFacePortExtractor1b.Inputs("CoordinateSystem").Add pCoordinateSystem
        'oFacePortExtractor1b.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort1")
        oFacePortExtractor1b.Parameter("GeometrySelector") = 4
        oFacePortExtractor1b.Parameter("FacesContext") = 2
        oFacePortExtractor1b.Parameter("LookingAxis") = 2
        oFacePortExtractor1b.Parameter("TrackFlag") = 2
        oFacePortExtractor1b.Parameter("Offset") = 0#
        oFacePortExtractor1b.Public = True
        oFacePortExtractor1b.Evaluate
        'Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort1"), oFacePortExtractor1b.ControlledInputs("Port"), sKeyOfMember1)
        
        Dim pPointAtSurfaceCenter1a As SP3DGeometricConstruction.GeometricConstruction
        Set pPointAtSurfaceCenter1a = pGCFactory.CreateEntity("PointAtSurfaceCenter", pPOM, "023d-PointAtSurfaceCenter")
        pPointAtSurfaceCenter1a.Inputs("Surface").Add oFacePortExtractor1a
        pPointAtSurfaceCenter1a.Evaluate
        
        Dim pPointAtSurfaceCenter1b As SP3DGeometricConstruction.GeometricConstruction
        Set pPointAtSurfaceCenter1b = pGCFactory.CreateEntity("PointAtSurfaceCenter", pPOM, "023d-PointAtSurfaceCenter")
        pPointAtSurfaceCenter1b.Inputs("Surface").Add oFacePortExtractor1b
        pPointAtSurfaceCenter1b.Evaluate
        
        Dim oParamDistMeasureParallel1 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel1 = pGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "023f-ParamDistMeasureParallel")
        oParamDistMeasureParallel1.Inputs("Graphics").Add pPointAtSurfaceCenter1a
        oParamDistMeasureParallel1.Inputs("Graphics").Add pPointAtSurfaceCenter1b
        oParamDistMeasureParallel1.Evaluate
        
        If CDbl(oParamDistMeasureParallel1.Parameter("Value")) > EPSILON Then
            Dim oSurfAtOffset1 As SP3DGeometricConstruction.GeometricConstruction
            Set oSurfAtOffset1 = pGCFactory.CreateEntity("SurfAtOffset", pPOM, "023g-SurfAtOffset")
            oSurfAtOffset1.Inputs("Surface").Add oFacePortExtractor1a
            oSurfAtOffset1.Inputs("TrackPoint").Add pPointAtSurfaceCenter1b
            oSurfAtOffset1.Parameter("Offset") = CDbl(oParamDistMeasureParallel1.Parameter("Value")) / 2
            oSurfAtOffset1.Evaluate
            
            Set oFacePortExtractor1a = oSurfAtOffset1
        End If
    End If
    
    If CInt(pGeometricConstruction.Parameter("Side")) <> 3 Then
        Dim oFacePortExtractor2a As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor2a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "024a-FacePortExtractor")
        oFacePortExtractor2a.Inputs("Connectable").Add oMember2
        oFacePortExtractor2a.Inputs("CoordinateSystem").Add pCoordinateSystem
        oFacePortExtractor2a.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort2")
        oFacePortExtractor2a.Parameter("GeometrySelector") = 4
        oFacePortExtractor2a.Parameter("FacesContext") = 2
        oFacePortExtractor2a.Parameter("LookingAxis") = 2
        oFacePortExtractor2a.Parameter("TrackFlag") = CInt(pGeometricConstruction.Parameter("Side"))
        oFacePortExtractor2a.Parameter("Offset") = 0#
        oFacePortExtractor2a.Public = True
        oFacePortExtractor2a.Evaluate
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort2"), oFacePortExtractor2a.ControlledInputs("Port"), sKeyOfMember2)
    Else
        Set oFacePortExtractor2a = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "024a-FacePortExtractor")
        oFacePortExtractor2a.Inputs("Connectable").Add oMember2
        oFacePortExtractor2a.Inputs("CoordinateSystem").Add pCoordinateSystem
        oFacePortExtractor2a.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort2")
        oFacePortExtractor2a.Parameter("GeometrySelector") = 4
        oFacePortExtractor2a.Parameter("FacesContext") = 2
        oFacePortExtractor2a.Parameter("LookingAxis") = 2
        oFacePortExtractor2a.Parameter("TrackFlag") = 1
        oFacePortExtractor2a.Parameter("Offset") = 0#
        oFacePortExtractor2a.Public = True
        oFacePortExtractor2a.Evaluate
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort2"), oFacePortExtractor2a.ControlledInputs("Port"), sKeyOfMember2)
        
        Dim oFacePortExtractor2b As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor2b = pGCFactory.CreateEntity("FacePortExtractor", pPOM, "024b-FacePortExtractor")
        oFacePortExtractor2b.Inputs("Connectable").Add oMember2
        oFacePortExtractor2b.Inputs("CoordinateSystem").Add pCoordinateSystem
        'oFacePortExtractor2b.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort2")
        oFacePortExtractor2b.Parameter("GeometrySelector") = 4
        oFacePortExtractor2b.Parameter("FacesContext") = 2
        oFacePortExtractor2b.Parameter("LookingAxis") = 2
        oFacePortExtractor2b.Parameter("TrackFlag") = 2
        oFacePortExtractor2b.Parameter("Offset") = 0#
        oFacePortExtractor2b.Public = True
        oFacePortExtractor2b.Evaluate
        'Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort2"), oFacePortExtractor2b.ControlledInputs("Port"), sKeyOfMember2)
    
        Dim pPointAtSurfaceCenter2a As SP3DGeometricConstruction.GeometricConstruction
        Set pPointAtSurfaceCenter2a = pGCFactory.CreateEntity("PointAtSurfaceCenter", pPOM, "023d-PointAtSurfaceCenter")
        pPointAtSurfaceCenter2a.Inputs("Surface").Add oFacePortExtractor2a
        pPointAtSurfaceCenter2a.Evaluate
        
        Dim pPointAtSurfaceCenter2b As SP3DGeometricConstruction.GeometricConstruction
        Set pPointAtSurfaceCenter2b = pGCFactory.CreateEntity("PointAtSurfaceCenter", pPOM, "023d-PointAtSurfaceCenter")
        pPointAtSurfaceCenter2b.Inputs("Surface").Add oFacePortExtractor2b
        pPointAtSurfaceCenter2b.Evaluate
            
        Dim oParamDistMeasureParallel2 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel2 = pGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "024f-ParamDistMeasureParallel")
        oParamDistMeasureParallel2.Inputs("Graphics").Add pPointAtSurfaceCenter2a
        oParamDistMeasureParallel2.Inputs("Graphics").Add pPointAtSurfaceCenter2b
        oParamDistMeasureParallel2.Evaluate
        
        If CDbl(oParamDistMeasureParallel2.Parameter("Value")) > EPSILON Then
            Dim oSurfAtOffset2 As SP3DGeometricConstruction.GeometricConstruction
            Set oSurfAtOffset2 = pGCFactory.CreateEntity("SurfAtOffset", pPOM, "024g-SurfAtOffset")
            oSurfAtOffset2.Inputs("Surface").Add oFacePortExtractor2a
            oSurfAtOffset2.Inputs("TrackPoint").Add pPointAtSurfaceCenter2b
            oSurfAtOffset2.Parameter("Offset") = CDbl(oParamDistMeasureParallel2.Parameter("Value")) / 2
            oSurfAtOffset2.Evaluate
            
            Set oFacePortExtractor2a = oSurfAtOffset2
        End If
    End If
    
    Dim oCurveByIntersection1 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection1 = pGCFactory.CreateEntity("CurveByIntersection", pPOM, "025-CurveByIntersection")
    oCurveByIntersection1.Inputs("Surface1").Add oTrimmingPlane1
    oCurveByIntersection1.Inputs("Surface2").Add oFacePortExtractor1a
    oCurveByIntersection1.Inputs("CoordinateSystem").Add pCoordinateSystem
    oCurveByIntersection1.Parameter("TrackFlag") = 1
    oCurveByIntersection1.Public = True
    oCurveByIntersection1.Evaluate

    Dim oCurveByIntersection2 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection2 = pGCFactory.CreateEntity("CurveByIntersection", pPOM, "026-CurveByIntersection")
    oCurveByIntersection2.Inputs("Surface1").Add oTrimmingPlane2
    oCurveByIntersection2.Inputs("Surface2").Add oFacePortExtractor2a
    oCurveByIntersection2.Inputs("CoordinateSystem").Add pCoordinateSystem
    oCurveByIntersection2.Parameter("TrackFlag") = 1
    oCurveByIntersection2.Public = True
    oCurveByIntersection2.Evaluate

    Dim oExtRuledSurface5 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtRuledSurface5 = pGCFactory.CreateEntity("ExtRuledSurface", pPOM, "027-ExtRuledSurface")
    oExtRuledSurface5.Inputs("Curve1").Add oCurveByIntersection1
    oExtRuledSurface5.Inputs("Curve2").Add oCurveByIntersection2
    oExtRuledSurface5.Parameter("Extension") = CDbl(pGeometricConstruction.Parameter("Extension"))
    oExtRuledSurface5.Public = True
    oExtRuledSurface5.Evaluate

    ' orientate the support
    If True Then
        ' extract z axis of the coordinate system
        Dim oLineFromCS As GeometricConstruction
        Set oLineFromCS = pGCFactory.CreateEntity("LineFromCS", pPOM, "028-LineFromCS")
        oLineFromCS.Input("CoordinateSystem") = pCoordinateSystem
        oLineFromCS.Parameter("LookingAxis") = 2#
        oLineFromCS.Parameter("Length") = 1#
        oLineFromCS.Public = True
        oLineFromCS.Evaluate
        
        Call SurfaceBody_OrientateFollowingLine(oExtRuledSurface5.Output("Support", 1), oLineFromCS, CInt(pGeometricConstruction.Parameter("Side")) = 1)
    End If
    
    ' generate the outputs
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGeometricConstruction
    pGeometricConstructionMacro.Output("Support", 1) = oExtRuledSurface5.Output("Support", 1)
    pGeometricConstructionMacro.Output("Boundary", "TR1." + sKeyOfMember1) = oTrimmingPlane1
    pGeometricConstructionMacro.Output("Boundary", "TR2." + sKeyOfMember2) = oTrimmingPlane2
    pGeometricConstructionMacro.Output("Boundary", "FacePort3") = oFacePortExtractor3a.Output
    pGeometricConstructionMacro.Output("Boundary", "FacePort4") = oFacePortExtractor4a.Output
    Call GeometricConstructionMacro_CreateCoordinateSystemOfSupport(pGeometricConstructionMacro, pPOM, GetGCGeomFactory(), GetGCGeomFactory2(), pCoordinateSystem.Origin)
    Call GeometricConstructionMacro_CreateNode(pGeometricConstructionMacro, pCoordinateSystem.Origin)
    
    Call DebugOut
End Sub
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
    
    ' loop on all the inputs, whose name is prefixed by "MemberPart"
    Dim pGeometricConstructionDefinition As IJGeometricConstructionDefinition: Set pGeometricConstructionDefinition = MyGC.definition
    Dim i  As Integer
    For i = 1 To pGeometricConstructionDefinition.InputCount
        Dim sName As String, sPrompt As String, sFilter As String, lMinConnected As Long, lMaxConnected As Long, sComputeIIDs As String
        Call pGeometricConstructionDefinition.GetInputInfoByIndex(i, sName, sPrompt, sFilter, lMinConnected, lMaxConnected, sComputeIIDs)
                
        ' migrate these inputs, if one of them has been replaced
        If Mid(sName, 1, Len("MemberPart")) = "MemberPart" Then Call GeometricConstruction_MigrateInputs(MyGC, pMigrateHelper, sName, pPositionOfNode)
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


