VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RingPlate3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const m_sModule = "GCSidePlate.RingPlate"
Private Enum Errors
    MISSING_MANDATORY_INPUT1 = 1
    MISSING_MANDATORY_INPUT2 = 2
    MISSING_MANDATORY_INPUT3 = 3
    COLUMN_NOT_CIRCULAR = 4
    COMMON_NODE_NOT_FOUND = 5
    MEMBER_AXIS_TOO_SMALL = 6
    MEMBER_AXES_NOT_COPLANAR = 7
    MEMBER_AXES_NOT_ORTHOGONAL = 8
    MEMBER_FACES_NOT_PARALLEL = 9
    MEMBER_NOT_CENTERED = 10
End Enum

Private Const sOrientation As String = "Orientation"

' define factories
Dim m_pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory

Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGeometricConstructionDynamic_IsOnRibbonBar
Implements IJGCMirror
Implements IJGCToDoDelegate
Implements IJGCConnectedElementSC

Private Sub Class_Initialize()
    Set m_pGeometricConstructionFactory = New GeometricConstructionEntitiesFactory
'    Let m_lCountOfDebugSupport = m_lCountOfDebugSupport + 1
'    Set m_pDebugSupport(m_lCountOfDebugSupport) = New DebugSupport
'    Let m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_SOURCE = "RingPlate (" + CStr(m_lCountOfDebugSupport) + ")"
'    Call m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_MSG(">> RingPlate")
End Sub
Private Sub Class_Terminate()
    Set m_pGeometricConstructionFactory = Nothing
'    Call m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_MSG("<< RingPlate")
'    Set m_pDebugSupport(m_lCountOfDebugSupport) = Nothing
'    Let m_lCountOfDebugSupport = m_lCountOfDebugSupport - 1
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
        Dim pElementsOfPrimaryMembers As IJElements: Set pElementsOfPrimaryMembers = pGC.Inputs("MemberParts")
        Dim iCountOfPrimaryMembers As Integer: iCountOfPrimaryMembers = pElementsOfPrimaryMembers.Count

        ' retrieve info about the secondary members
        Dim pElementsOfSecondaryMembers As IJElements: Set pElementsOfSecondaryMembers = pGC.Inputs("SecondaryMemberParts")
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
    ' located inputs
    Call pGeometricConstructionDefinition.AddInput("MemberParts", "Select n coplanar Members in consecutive counterclockwise or clockwise order", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 1, 100, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("OrthogonalMemberPart", "Select 1 orthogonal Member", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SecondaryMemberParts", "Select n coplanar secondary Members", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 0, 8, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")

    ' controlled inputs
    Call pGeometricConstructionDefinition.AddControlledInput("AdvancedPlateSystem", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort1") ' relative to the first smart step
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1") ' relative to the first smart step
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort2") ' relative to the second smart step
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort2") ' relative to the second smart step
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort3") ' relative to the third smart step
    
    ' special controlled input to indicate that some PlateSystems of incoming Members have not been correctly un-bounded during the disconnect semantic,
    ' in the case, where the APS and some incoming Members are simultaneously deleted.
    ' this is due to the fact that FrameConnections depending from a deleted Member are first deleted, before being later recreated.
    ' our disconnect semantic executes, after those FrameConnections get deleted and before they get re-created.
    ' we have to wait the triggering of our compute semantic to have access again to the re-created ones.
    Call pGeometricConstructionDefinition.AddControlledInput("UnBoundedPlateSystems", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    ' parameters
    Call pGeometricConstructionDefinition.AddParameter("WeldToe", "WeldToe", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1.5)
    Call pGeometricConstructionDefinition.AddParameter("Fillet", "Fillet", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("CutBack", "CutBack", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Extension", "Extension", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    
    ' code listed parameters
    Call pGeometricConstructionDefinition.AddParameter("Support", "Support", GCCodeList, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Back", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Front", 2)
    
    ' parameter for cs orientation (added for mirror)
    Call pGeometricConstructionDefinition.AddParameter(sOrientation, "Orientation", GCCodeList, 0, 0, 0, 0, GCDirect)
    Call pGeometricConstructionDefinition.AddParameterValue(sOrientation, "Direct", GCDirect)
    Call pGeometricConstructionDefinition.AddParameterValue(sOrientation, "Indirect", GCIndirect)
    
    Call pGeometricConstructionDefinition.AddParameter("IsCutByTube", "IsCutByTube", GCCodeList, 0, 0, 0, 0, 1, True)
    Call pGeometricConstructionDefinition.AddParameterValue("IsCutByTube", "False", 0)
    Call pGeometricConstructionDefinition.AddParameterValue("IsCutByTube", "True", 1)
    
    ' error codes
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MANDATORY_INPUT1, "MissingMandatoryInput1", "At least 1 member are required")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MANDATORY_INPUT2, "MissingMandatoryInput2", "An orthogonal member is required")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MANDATORY_INPUT3, "MissingMandatoryInput3", "Only 1 orthogonal member is required")
    Call pGeometricConstructionDefinition.AddErrorValue(COLUMN_NOT_CIRCULAR, "ColumnNotCircular", "The orthogonal member is not circular")
    Call pGeometricConstructionDefinition.AddErrorValue(COMMON_NODE_NOT_FOUND, "CommonNodeNotFound", "Common node not found")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXIS_TOO_SMALL, "MemberAxisTooSmall", "At least 1 member axis is too small")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXES_NOT_COPLANAR, "MemberAxesNotCoplanar", "Member axes not coplanar")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXES_NOT_ORTHOGONAL, "MemberAxesNotOrthogonal", "Member axes not perpendicular to the orthogonal member")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_FACES_NOT_PARALLEL, "MemberFacesNotParallel", "Member faces not parallel")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_NOT_CENTERED, "MemberNotCentered", "Member is not Centered, Top-Centered or Bottom-Centered")
    
    ' outputs
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Support")
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypePoint3d, "PseudoBoundary")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypePoint3d, "Node")
    Call pGeometricConstructionDefinition.AddOutput(GCLocalCoordinateSystem, "CoordinateSystem")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    
    ' use GC as a GCmacro
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGeometricConstruction
    
    ' retrieve info about the incoming members
    Dim pElementsOfMembers As IJElements: Set pElementsOfMembers = pGeometricConstruction.Inputs("MemberParts")
    Dim iCountOfMembers As Integer: Let iCountOfMembers = pElementsOfMembers.Count
    Dim sKeysOfMembers() As String: If iCountOfMembers > 0 Then Let sKeysOfMembers = Elements_GetKeys(pElementsOfMembers)
    If iCountOfMembers < 1 Then Err.Raise MISSING_MANDATORY_INPUT1
    
    ' retrieve info about the orthogonal member
    Dim oOrthogonalMember As Object
    If True Then
        ' check for errors on orthogonal members
        If pGeometricConstruction.Inputs("OrthogonalMemberPart").Count = 0 Then Err.Raise MISSING_MANDATORY_INPUT2
        If pGeometricConstruction.Inputs("OrthogonalMemberPart").Count > 1 Then Err.Raise MISSING_MANDATORY_INPUT3
        
        ' get the orthogonal member
        Set oOrthogonalMember = pGeometricConstruction.Input("OrthogonalMemberPart")
    End If
       
    ' retrieve info about the secondary members
    Dim pElementsOfSecondaryMembers As IJElements: Set pElementsOfSecondaryMembers = pGeometricConstruction.Inputs("SecondaryMemberParts")
    Dim iCountOfSecondaryMembers As Integer: Let iCountOfSecondaryMembers = pElementsOfSecondaryMembers.Count
    Dim sKeysOfSecondaryMembers() As String: If iCountOfSecondaryMembers > 0 Then Let sKeysOfSecondaryMembers = Elements_GetKeys(pElementsOfSecondaryMembers)
    
    ' retrieve info about the boundaries
    Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
    Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
    Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)
        
    ' retrieve info about the pseudo-boundaries
    Dim pElementsOfPseudoBoundaries As IJElements: Set pElementsOfPseudoBoundaries = pGeometricConstructionMacro.Outputs("PseudoBoundary")
    Dim iCountOfPseudoBoundaries As Integer: Let iCountOfPseudoBoundaries = pElementsOfPseudoBoundaries.Count
    Dim sKeysOfPseudoBoundaries() As String: If iCountOfPseudoBoundaries > 0 Then Let sKeysOfPseudoBoundaries = Elements_GetKeys(pElementsOfPseudoBoundaries)
    
    ' retrieve lines of members axes
    Dim pLinesOfMemberAxes() As IJLine: Let pLinesOfMemberAxes = Members_GetLinesOfMemberPartAxes(pElementsOfMembers)
    Dim pLineOfOrthogonalMemberAxis As IJLine: Set pLineOfOrthogonalMemberAxis = Member_GetLine(oOrthogonalMember)
    If Lines_GetMinimumLength(pLinesOfMemberAxes) < EPSILON _
    Or pLineOfOrthogonalMemberAxis.Length() < EPSILON Then Err.Raise MEMBER_AXIS_TOO_SMALL
    
    ' verify the cardinal point of the incoming Members
    If Not IsCardinalPointOfMembersValid(pElementsOfMembers, iCountOfMembers) Then Err.Raise MEMBER_NOT_CENTERED
    
    ' compute or retrieve common node
    Dim pPositionOfNode As IJDPosition: Set pPositionOfNode = ComputePositionOfNode(pGeometricConstructionMacro, pLinesOfMemberAxes, iCountOfMembers, pLineOfOrthogonalMemberAxis)
    If pPositionOfNode Is Nothing Then Err.Raise COMMON_NODE_NOT_FOUND
   
    ' compute support and boundaries - the array of keys for boundaries is passed ByRef and will be updated within the sub
    Call CreateNestedMacros(pPOM, pGeometricConstruction, _
                            pLinesOfMemberAxes, pPositionOfNode, _
                            pLineOfOrthogonalMemberAxis, _
                            iCountOfMembers, pElementsOfMembers, sKeysOfMembers, _
                            iCountOfBoundaries, pElementsOfBoundaries, sKeysOfBoundaries)
            
    Call CreatePseudoBoundaries(pPOM, pGeometricConstruction, _
                                iCountOfSecondaryMembers, pElementsOfSecondaryMembers, sKeysOfSecondaryMembers, _
                                iCountOfPseudoBoundaries, pElementsOfPseudoBoundaries, sKeysOfPseudoBoundaries)
    
    ' if the APS is connected, then we can un-bound or bound added Members or modify the Plate to bound, if the context has changed (modify Support)
    If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
         ' retrieve advanced plate system
        Dim oPlateSystemAdvanced As Object: Set oPlateSystemAdvanced = pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1)

        ' add new boundaries
        ' add new boundaries
        Call PlateSystem_AddNewBoundaries(oPlateSystemAdvanced, pGeometricConstructionMacro)
       
       ' get the key of the APS
        Dim sKey As String: Let sKey = pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").GetKey(oPlateSystemAdvanced)
        
        ' if the key is not set to "NoBound" then do the bound (this means that the APS has already computed its geometry and edge ports can be extracted)
        If Not sKey = "NotBounded" Then
            ' bound/un-bound members
            Call Members_BoundByOrUnboundFromEdgeOfAPS(pElementsOfMembers, iCountOfMembers, sKeysOfMembers, _
                                                       pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                                       pGeometricConstruction, 1, oPlateSystemAdvanced, pPositionOfNode)
            
            Call Members_BoundByOrUnboundFromFaceOfAPS(pElementsOfSecondaryMembers, oPlateSystemAdvanced, pPositionOfNode)
        End If
    End If
    
    ' remove unneeded boundaries
    If iCountOfBoundaries > 0 Then Call Elements_RemoveUnneeded(pElementsOfBoundaries, sKeysOfBoundaries)
    If iCountOfPseudoBoundaries > 0 Then Call Elements_RemoveUnneeded(pElementsOfPseudoBoundaries, sKeysOfPseudoBoundaries)
    
    Exit Sub
ErrorHandler:
    Call GCProcessError(pGeometricConstruction, , Err.Number)
End Sub
'
' implementation of the IJGCSemanticConnection interface
'
Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal oRelationship As Object)
    Dim pRelationship As IJDRelationship: Set pRelationship = oRelationship
'    MsgBox "IJGCSemanticConnection_PostConnectionAdded: " + pRelationship.Name
End Sub
Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
    Dim pRelationship As IJDRelationship: Set pRelationship = oRelationship
    Dim pGeometricConstruction As IJGeometricConstruction: Set pGeometricConstruction = pRelationship.Origin
    
    If Mid(pRelationship.Name, 1, Len("AdvancedPlateSystem")) = "AdvancedPlateSystem" And bIsDestinationDeleted Then
        ' the APS is deleting
        ' unbound remaining connected Members
        On Error Resume Next ' avoid failure to not prevent the APS to get deleted
        Call Members_UnboundOnDelete(pGeometricConstruction, pRelationship.Destination)
        On Error GoTo 0
        
        ' delete the GCMacro itself, if no more processing is expected from the compute semantic
        If pGeometricConstruction.ControlledInputs("UnBoundedPlateSystems").Count = 0 Then Call Object_Delete(pGeometricConstruction)
    ElseIf Mid(pRelationship.Name, 1, Len("MemberParts")) = "MemberParts" And bIsDestinationDeleted Then
        ' an input member is deleting, disconnect the corresponding ports
        Dim sKey As String: Let sKey = Mid(pRelationship.Name, Len("MemberParts") + 2)
        pGeometricConstruction.ControlledInputs("AxisPort1").Remove (sKey)
        pGeometricConstruction.ControlledInputs("FacePort1").Remove (sKey)
    ElseIf Mid(pRelationship.Name, 1, Len("MemberParts")) = "MemberParts" Or Mid(pRelationship.Name, 1, Len("SecondaryMemberParts")) = "SecondaryMemberParts" Then
         ' Unbound removed member if the APS is connected
        If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
            Call Member_UnboundOnRemove(pGeometricConstruction, pRelationship.Destination)
        Else
            'input is removed clear controlled inputs if GC not deleted
            If bIsOriginDeleted = False Then
                pGeometricConstruction.ControlledInputs("AxisPort1").Clear
                pGeometricConstruction.ControlledInputs("AxisPort2").Clear
                pGeometricConstruction.ControlledInputs("AxisPort3").Clear
                pGeometricConstruction.ControlledInputs("FacePort1").Clear
                pGeometricConstruction.ControlledInputs("FacePort2").Clear
            End If
        End If
    End If
    
End Sub
'
' private APIs
'
Private Sub CreatePseudoBoundaries(pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                   iCountOfSecondaryMembers As Integer, pElementsOfSecondaryMembers As IJElements, sKeysOfSecondaryMembers() As String, _
                                   iCountOfBoundaries As Integer, pElementsOfBoundaries As IJElements, ByRef sKeysOfBoundaries() As String)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "CreatePseudoBoundaries"
    
    ' retrieve GCMacro
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGeometricConstruction
    
    ' instantiate GCFactory
    Dim pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory: Set pGeometricConstructionFactory = New GeometricConstructionEntitiesFactory
    
    ' compute dummy position
    Dim pPosition As IJDPosition
    If True Then
        Set pPosition = New DPosition
        Call pPosition.Set(0#, 0#, 0#)
    End If
    
    Let sComment = "loop on secondary members"
    Dim i As Integer
    For i = 1 To iCountOfSecondaryMembers
        ' set controlled input
        If True Then
            ' extract the axis port
            Dim oAxisPort As Object
            If True Then
                Dim oLineAxisPortExtractor0 As GeometricConstruction
                Set oLineAxisPortExtractor0 = pGeometricConstructionFactory.CreateEntity("LineAxisPortExtractor", pPOM)
                oLineAxisPortExtractor0.Inputs("MemberPart").Add pElementsOfSecondaryMembers(i)
                oLineAxisPortExtractor0.Evaluate
                Set oAxisPort = oLineAxisPortExtractor0.ControlledInputs("Port")(1)
            End If
            
            ' update the corresponding controlled input
            If True Then
                Dim pElementsOfAxisPorts As IJElements: Set pElementsOfAxisPorts = pGeometricConstruction.ControlledInputs("AxisPort3")
                If Not pElementsOfAxisPorts.Contains(oAxisPort) Then
                    Call pElementsOfAxisPorts.Add(oAxisPort, sKeysOfSecondaryMembers(i))
                End If
            End If
        End If
        
        ' build key
        Dim sKeyOfPseudoBoundary As String: Let sKeyOfPseudoBoundary = "PB3." + sKeysOfSecondaryMembers(i)
        ' assign a dummy position to each pseudo boundary
        pGeometricConstructionMacro.Output("PseudoBoundary", sKeyOfPseudoBoundary) = Point_FromPosition(pPosition)
    
         If iCountOfBoundaries > 0 Then
            Call Keys_RemoveKey(sKeysOfBoundaries, sKeyOfPseudoBoundary)
        End If
    Next
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Private Sub Members_UnboundOnDelete(pGeometricConstructionMacro As IJGeometricConstructionMacro, oAdvancedPlateSystem As Object)
    ' use the GCMacro as a GC
    Dim pGeometricConstruction As IJGeometricConstruction: Set pGeometricConstruction = pGeometricConstructionMacro
    
    ' retrieve info about the incoming members
    Dim pElementsOfMembers As IJElements: Set pElementsOfMembers = Elements_GetValidElements(pGeometricConstruction.Inputs("MemberParts"))
    Dim iCountOfMembers As Integer: Let iCountOfMembers = pElementsOfMembers.Count
    Dim sKeysOfMembers() As String: If iCountOfMembers > 0 Then Let sKeysOfMembers = Elements_GetDummyKeys(pElementsOfMembers)
        
    ' retrieve info about the orthogonal member
    Dim oOrthogonalMember As Object
    If True Then
        Dim pElementsOfValidElements As IJElements:
        Set pElementsOfValidElements = Elements_GetValidElements(pGeometricConstruction.Inputs("OrthogonalMemberPart"))
        If pElementsOfValidElements.Count = 1 Then Set oOrthogonalMember = pElementsOfValidElements(1)
    End If
    
    ' retrieve position of node
    If iCountOfMembers > 0 Then
        ' compute the node
        Dim pPositionOfNode As IJDPosition
        If Not oOrthogonalMember Is Nothing Then
            ' if something goes wrong, we will at least try to retrieve the node
            On Error Resume Next
        
            ' retrieve axes of incoming members
            Dim pLinesOfMemberAxes() As IJLine: Let pLinesOfMemberAxes = Members_GetLinesOfMemberPartAxes(pElementsOfMembers)
            
            ' retrieve axis of orthogonal member
            Dim pLineOfOrthogonalMemberAxis As IJLine
            If Not oOrthogonalMember Is Nothing Then Set pLineOfOrthogonalMemberAxis = Member_GetLine(oOrthogonalMember)
            
            If Lines_GetMinimumLength(pLinesOfMemberAxes) > EPSILON _
            And pLineOfOrthogonalMemberAxis.Length() > EPSILON Then
                ' compute common node
                Set pPositionOfNode = ComputePositionOfNode(pGeometricConstructionMacro, pLinesOfMemberAxes, iCountOfMembers, pLineOfOrthogonalMemberAxis)
            End If
            
            On Error GoTo 0
        End If
        
        ' retrieve the node, if not computed
        If pPositionOfNode Is Nothing Then Set pPositionOfNode = GeometricConstructionMacro_RetrievePositionOfNode(pGeometricConstruction)
    End If
    
    If iCountOfMembers > 0 Then
        ' retrieve info about the boundaries
        Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
        Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
        Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)
        


        ' un-bound removed members
        Call Members_UnboundFromEdgeOfAPS(pElementsOfMembers, iCountOfMembers, sKeysOfMembers, _
                                          pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                          pGeometricConstruction, 1, oAdvancedPlateSystem, pPositionOfNode)
    End If
    
    ' retrieve info about the secondary members
    Dim pElementsOfSecondaryMembers As IJElements: Set pElementsOfSecondaryMembers = Elements_GetValidElements(pGeometricConstruction.Inputs("SecondaryMemberParts"))
    Dim iCountOfSecondaryMembers As Integer: Let iCountOfSecondaryMembers = pElementsOfSecondaryMembers.Count
    Dim sKeysOfSecondaryMembers() As String: If iCountOfSecondaryMembers > 0 Then Let sKeysOfSecondaryMembers = Elements_GetDummyKeys(pElementsOfSecondaryMembers)

    If iCountOfSecondaryMembers > 0 Then
        ' retrieve info about the boundaries
        Dim pElementsOfPseudoBoundaries As IJElements: Set pElementsOfPseudoBoundaries = pGeometricConstructionMacro.Outputs("PseudoBoundary")
        Dim iCountOfPseudoBoundaries As Integer: Let iCountOfPseudoBoundaries = pElementsOfPseudoBoundaries.Count
        Dim sKeysOfPseudoBoundaries() As String: If iCountOfPseudoBoundaries > 0 Then Let sKeysOfPseudoBoundaries = Elements_GetKeys(pElementsOfPseudoBoundaries)

        ' un-bound removed secondary members
        Call Members_UnboundFromFaceOfAPS(pElementsOfSecondaryMembers, iCountOfSecondaryMembers, sKeysOfSecondaryMembers, _
                                          pElementsOfPseudoBoundaries, iCountOfPseudoBoundaries, sKeysOfPseudoBoundaries, _
                                          pGeometricConstruction, 3, oAdvancedPlateSystem, pPositionOfNode)

    End If
End Sub
Private Function ComputePositionOfNode(pGeometricConstructionMacro As IJGeometricConstructionMacro, _
                                       pLinesOfMemberAxes() As IJLine, iCountOfMembers As Integer, _
                                       pLineOfOrthogonalMemberAxis As IJLine) As IJDPosition
    ' prepare result
    Dim pPositionOfNode As IJDPosition
    
    ' compute common point on curve
    Set pPositionOfNode = GetPositionAtExtremityOfLineOnCurve(pLinesOfMemberAxes(1), pLineOfOrthogonalMemberAxis)
    
    ' verify that common node also works for regular members
    If Not pPositionOfNode Is Nothing And iCountOfMembers > 1 Then
        Dim pPositionOfNode1 As IJDPosition: Set pPositionOfNode1 = GetPositionAtCommonExtremityOfLines(pLinesOfMemberAxes)
        If pPositionOfNode1 Is Nothing Then
            Set pPositionOfNode = Nothing
        Else
            If pPositionOfNode1.DistPt(pPositionOfNode) > EPSILON Then Set pPositionOfNode = Nothing
        End If
    End If
    
    ' return result
    Set ComputePositionOfNode = pPositionOfNode
End Function
Private Function IsCardinalPointOfMembersValid(pElementsOfMembers As IJElements, iCountOfMembers As Integer) As Boolean
    ' prepare result
    Dim bIsCardinalPointOfMembersValid As Boolean: Let bIsCardinalPointOfMembersValid = True
    
    Dim i As Integer
    For i = 1 To iCountOfMembers
        Dim oMember As Object
        Set oMember = pElementsOfMembers(i)
        Dim pCrossSection As ISPSCrossSection
        If TypeOf oMember Is ISPSMemberPartPrismatic Then
            Dim pMemberPartPrismatic As ISPSMemberPartPrismatic
            Set pMemberPartPrismatic = oMember
            Set pCrossSection = pMemberPartPrismatic.CrossSection
        ElseIf TypeOf oMember Is ISPSDesignedMember Then
            Dim pDesignedMember As ISPSDesignedMember
            Set pDesignedMember = oMember
            Set pCrossSection = pDesignedMember
        Else
            MsgBox "Invalid StructProfilePart"
        End If

        If pCrossSection.CardinalPoint <> 5 _
        And pCrossSection.CardinalPoint <> 2 _
        And pCrossSection.CardinalPoint <> 8 Then
            Let bIsCardinalPointOfMembersValid = False
            Exit For
        End If
    Next
    
    ' return result
    Let IsCardinalPointOfMembersValid = bIsCardinalPointOfMembersValid
End Function
Private Sub CreateNestedMacros(pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                       pLinesOfMemberAxes() As IJLine, pPositionOfNode As IJDPosition, _
                       pLineOfOrthogonalLine As IJLine, _
                       iCountOfMembers As Integer, pElementsOfMembers As IJElements, sKeysOfMembers() As String, _
                       iCountOfBoundaries As Integer, pElementsOfBoundaries As IJElements, ByRef sKeysOfBoundaries() As String)
    ' check if orthogonal member is orthogonal
    If Not AreVectorsPerpendicular(Vector_FromLine(pLinesOfMemberAxes(1)), Vector_FromLine(pLineOfOrthogonalLine)) Then Err.Raise MEMBER_AXES_NOT_ORTHOGONAL

    ' compute line normal to lines
    Dim pLineOfNormal As IJLine: Set pLineOfNormal = GetLineAtNormalToLinesAndParallelToLine(GetGCGeomFactory(), GetGCGeomFactory2(), pPOM, pLinesOfMemberAxes, pLineOfOrthogonalLine)
    If pLineOfNormal Is Nothing Then Err.Raise MEMBER_AXES_NOT_COPLANAR
    
    ' order lines
    Dim lOrderedIndexes() As Long
    Dim dOrderedAngles() As Double
    Call Lines_OrderAroundPointAndNormal(pLinesOfMemberAxes, pPositionOfNode, Vector_FromLine(pLineOfNormal), lOrderedIndexes, dOrderedAngles)
        
    ' create macros "ExtractPorts"
    Dim pGeometricConstructionMacrosOfExtractPorts() As IJGeometricConstructionMacro: ReDim pGeometricConstructionMacrosOfExtractPorts(1 To iCountOfMembers)
    Dim iMember As Integer
    For iMember = 1 To iCountOfMembers
        Set pGeometricConstructionMacrosOfExtractPorts(iMember) = CreateGeometricConstructionMacroOfExtractPorts(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                                 pElementsOfMembers(iMember), sKeysOfMembers(iMember), _
                                                                                                                 pLineOfNormal)
        ' check if faces are coplanar
        If iMember > 1 Then If Not ArePlanesParallel(pGeometricConstructionMacrosOfExtractPorts(iMember).Outputs("Support").Item(1), _
                                                     pGeometricConstructionMacrosOfExtractPorts(1).Outputs("Support").Item(1)) Then _
            Err.Raise MEMBER_FACES_NOT_PARALLEL
    Next
   
    ' populate output "Node"
    Call GeometricConstructionMacro_CreateNode(pGeometricConstruction, pPositionOfNode)
    
    ' populate output "Support"
    Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Support", 1, pGeometricConstructionMacrosOfExtractPorts(1).Outputs("Support")(1))
       
    ' compute output "CoordinateSystem"
    Call GeometricConstructionMacro_CreateCoordinateSystemOfSupport(pGeometricConstruction, pPOM, GetGCGeomFactory(), GetGCGeomFactory2(), pPositionOfNode)
    
    ' populate macro "ExtractCircle"
    Dim pGeometricConstructionMacroOfExtractCircle As IJGeometricConstructionMacro
    Set pGeometricConstructionMacroOfExtractCircle = CreateGeometricConstructionMacroOfExtractCircle(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction)
    
    ' populate first output "Boundary"
    If CInt(pGeometricConstruction.Parameter("IsCutByTube")) = 1 Then
        Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Boundary", "FacePort2", pGeometricConstructionMacroOfExtractCircle.Outputs("Boundary")(1))
    Else
        Call GeometricConstructionMacro_DeleteOutput(pGeometricConstruction, "Boundary", "FacePort2")
    End If

    If iCountOfBoundaries > 0 Then
        Call Keys_RemoveKey(sKeysOfBoundaries, "FacePort2")
    End If
    
    ' create macros "FreeEdge"
    Dim pGeometricConstructionMacrosOfFreeEdges() As IJGeometricConstructionMacro: ReDim pGeometricConstructionMacrosOfFreeEdges(1 To iCountOfMembers)
    Dim iFreeEdge As Integer
    For iFreeEdge = 1 To iCountOfMembers
        ' index of the first Member supporting the FreeEdge
        Dim iMember1 As Integer: Let iMember1 = lOrderedIndexes(iFreeEdge)
        
        ' index of th second Member supporting the FreeEdge
        Dim iMember2 As Integer: If iFreeEdge < iCountOfMembers Then Let iMember2 = lOrderedIndexes(iFreeEdge + 1) Else Let iMember2 = lOrderedIndexes(1)
                
        ' angle between the 2 Member axes
        Dim dAngle As Double: If iFreeEdge < iCountOfMembers Then Let dAngle = dOrderedAngles(iFreeEdge + 1) - dOrderedAngles(iFreeEdge) Else Let dAngle = 360 - dOrderedAngles(iFreeEdge)

        ' compute "FreeEdge"
        Set pGeometricConstructionMacrosOfFreeEdges(iFreeEdge) = CreateGeometricConstructionMacroOfStraightEdgeAround3(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                                      pElementsOfMembers(iMember1), _
                                                                                                                      pGeometricConstructionMacrosOfExtractPorts(iMember1), _
                                                                                                                      pGeometricConstructionMacrosOfExtractPorts(iMember2), _
                                                                                                                      pGeometricConstructionMacroOfExtractCircle, _
                                                                                                                      pLineOfNormal, dAngle)
        
        ' populate output "Boundary"
        Dim sKeyOfFreeEdge As String: Let sKeyOfFreeEdge = "FE1." + sKeysOfMembers(iMember1) + "-" + sKeysOfMembers(iMember2)
        Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Boundary", sKeyOfFreeEdge, pGeometricConstructionMacrosOfFreeEdges(iFreeEdge).Outputs("Boundary")(1))

        If iCountOfBoundaries > 0 Then
            Call Keys_RemoveKey(sKeysOfBoundaries, sKeyOfFreeEdge)
        End If
    Next
    
    ' create macros "TrimmingPoint" and "TrimmingPlane"
    For iMember = 1 To iCountOfMembers
        ' index of the first FreeEdge used by the TrimmingPoint
        Dim iFreeEdge1 As Integer
        Dim j As Integer
        For j = 1 To iCountOfMembers:
            If lOrderedIndexes(j) = iMember Then
                Let iFreeEdge1 = j:
                Exit For:
            End If
        Next
        
        ' index of the second FreeEdge used by the TrimmingPoint
        Dim iFreeEdge2 As Integer: If iFreeEdge1 = 1 Then Let iFreeEdge2 = iCountOfMembers Else Let iFreeEdge2 = iFreeEdge1 - 1

        ' create macro "TrimmingPoint"
        Dim pGeometricConstructionMacroOfTrimmingPoint As IJGeometricConstructionMacro
        Set pGeometricConstructionMacroOfTrimmingPoint = CreateGeometricConstructionMacroOfTrimmingPoint(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                         pGeometricConstructionMacrosOfExtractPorts(iMember), _
                                                                                                         pGeometricConstructionMacrosOfFreeEdges(iFreeEdge1), _
                                                                                                         pGeometricConstructionMacrosOfFreeEdges(iFreeEdge2))
        ' create macro "TrimmingPlane"
        Dim pGeometricConstructionMacroOfTrimmingPlane As IJGeometricConstructionMacro
        Set pGeometricConstructionMacroOfTrimmingPlane = CreateGeometricConstructionMacroOfTrimmingPlane(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                         pElementsOfMembers(iMember), _
                                                                                                         pGeometricConstructionMacroOfTrimmingPoint)
        ' populate output "Boundary"
        Dim sKeyOfTrimmingPlane As String: Let sKeyOfTrimmingPlane = "TR1." + sKeysOfMembers(iMember)
        Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Boundary", sKeyOfTrimmingPlane, ModelBody_FromGeometry(pGeometricConstructionMacroOfTrimmingPlane.Outputs("Plane")(1)))
        
        If iCountOfBoundaries > 0 Then
            Call Keys_RemoveKey(sKeysOfBoundaries, sKeyOfTrimmingPlane)
        End If
    Next
End Sub

Private Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal info As Variant, IsOnRibbonBar As Boolean)
    IsOnRibbonBar = True

    Select Case sName
        Case sOrientation: IsOnRibbonBar = False
    End Select

End Sub
'*******************************************************************************************************************
 'Adaptation of the GC for mirror copy or copy symmetry operations.
 '   This sub is called by the GC evaluate method (CGeometricConstruction::Evaluate) when a mirror is in progress
 '   (CTL_FLAG_MIRROR_IN_PROGRESS flag is set on the BO during mirror operations (CGeometricConstruction::Adapt))
 '   At this stage, the GC output (aggregated geometry) has been processed by the mirror (mirrored /adapted or transformed)
 '   and holds the mirrored geometry that we assume being the exact mirror. The inputs of the mirrored GC are also well defined:
 '   mirrored if they were in the copy set, identified, kept as original, or deleted (delete optional) by the user in the adapt form.
 '
 '   The purpose of this adapt method is to adapt the GC parameters so that the next recompute will match this exact mirrored geometry
 '   or be as close as possible.
 '
 '   FAILURE:
 '   We assume here that any failure or exception is caught by the caller which is the GC BO.
 '
 '   Adaptation of GCRingPlate3 GC:
 '      Mirror z axis of the macro 'CoordinateSystem' output and reverse sOrientation' parameter otherwise
 '
Private Sub IJGCMirror_Adapt(ByVal pOriginalGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pMirroredGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pMirrorPlane As IngrGeom3D.IJPlane, ByVal pMatrix As AutoMath.IJDT4x4, ByVal pGCInputsFromInitialCopySet As IMSCoreCollections.IJElements)
    'First check if mirror parameter has been bulkloaded otherwise exit
    Dim pGCPrivateAccess As IJGCPrivateAccess: Set pGCPrivateAccess = pOriginalGC
    Dim pGCType As IJGeometricConstructionType: Set pGCType = pGCPrivateAccess.GeometricConstructionType
    If pGCType.IsParameterBulkloaded(sOrientation) = False Then
        Exit Sub
    End If
    
    'Get the original GCmacro 'CoordinateSystem'
    Dim pOriginalMacroCoordinateSystem  As IJLocalCoordinateSystem: Set pOriginalMacroCoordinateSystem = pOriginalGC.Output("CoordinateSystem", 1)
    
    Dim pPOMForDebug As IJDPOM
    If GetGlobalShowDetails() Then Set pPOMForDebug = Object_GetPOM(pOriginalMacroCoordinateSystem)
    
    'mirror original z axis
    Dim pMirroredZAxis As IJDVector: Set pMirroredZAxis = pMatrix.TransformVector(pOriginalMacroCoordinateSystem.ZAxis)

    'Reverse 'Support' parameter if Z axis is reversed by mirror
    'Otherwise reverse sOrientation flag that will be used by ExtractPorts
    pMirroredGC.Parameter("Support") = pOriginalGC.Parameter("Support")
    If pMirroredZAxis.Dot(pOriginalMacroCoordinateSystem.ZAxis) < -EPSILON Then
        Select Case pOriginalGC.Parameter("Support")
            Case 1:
                pMirroredGC.Parameter("Support") = 2
            Case 2:
                pMirroredGC.Parameter("Support") = 1
        End Select
    Else
        Select Case pOriginalGC.Parameter(sOrientation)
            Case GCDirect:
                pMirroredGC.Parameter(sOrientation) = GCIndirect
            Case GCIndirect:
                pMirroredGC.Parameter(sOrientation) = GCDirect
        End Select
    End If

End Sub

Private Function GetGlobalShowDetails() As Long
    Dim oToolshelper As New PORTHELPERLib.ToolsHelper
    GetGlobalShowDetails = oToolshelper.GetGlobalShowDetails()
End Function
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



