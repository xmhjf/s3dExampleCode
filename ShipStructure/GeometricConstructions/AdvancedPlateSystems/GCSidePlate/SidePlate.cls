VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SidePlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Enum Errors
    MISSING_MANDATORY_INPUT = 1
    COMMON_NODE_NOT_FOUND = 2
    MEMBER_AXIS_TOO_SMALL = 3
    MEMBER_AXES_COLINEAR = 4
    MEMBER_AXES_NOT_COPLANAR = 5
    MEMBER_FACES_NOT_PARALLEL = 6
End Enum

Private Const sMEMBER_PARTS As String = "MemberParts"
Private Const sSECONDARY_MEMBER_PARTS As String = "SecondaryMemberParts"
Private Const sWELD_TOE As String = "WeldToe"
Private Const sFILLET As String = "Fillet"
Private Const sARM_LENGTH As String = "ArmLength"
Private Const sCUT_BACK As String = "CutBack"
Private Const sEXTENSION As String = "Extension"
Private Const sSUPPORT As String = "Support"
Private Const sOrientation As String = "Orientation"

Const m_sModule = "GCSidePlate.SidePlate"

' define factories
Dim m_pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory

Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGCMigrate
Implements IJSidePlate
Implements IJGeometricConstructionDynamic_IsOnRibbonBar
Implements IJGCMirror
Implements IJGCToDoDelegate
Implements IJGCConnectedElementSC

Private Property Get Source() As String
    Let Source = "GCSidePlate.SidePlate"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property
Private Sub Class_Initialize()
    Call DebugStart(Source)

    Set m_pGeometricConstructionFactory = New GeometricConstructionEntitiesFactory
    '    Let m_lCountOfDebugSupport = m_lCountOfDebugSupport + 1
    '    Set m_pDebugSupport(m_lCountOfDebugSupport) = New DebugSupport
    '    Let m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_SOURCE = "SidePlate (" + CStr(m_lCountOfDebugSupport) + ")"
    '    Call m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_MSG(">> SidePlate")
End Sub
Private Sub Class_Terminate()
    Set m_pGeometricConstructionFactory = Nothing
    '    Call m_pDebugSupport(m_lCountOfDebugSupport).DEBUG_MSG("<< SidePlate")
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
        Dim pElementsOfPrimaryMembers As IJElements: Set pElementsOfPrimaryMembers = pGC.Inputs(sMEMBER_PARTS)
        Dim iCountOfPrimaryMembers As Integer: iCountOfPrimaryMembers = pElementsOfPrimaryMembers.Count

        ' retrieve info about the secondary members
        Dim pElementsOfSecondaryMembers As IJElements: Set pElementsOfSecondaryMembers = pGC.Inputs(sSECONDARY_MEMBER_PARTS)
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
    
    ' located inputs
    Call pGeometricConstructionDefinition.AddInput(sMEMBER_PARTS, "Select n coplanar Members in consecutive counterclockwise or clockwise order", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 2, 8, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput(sSECONDARY_MEMBER_PARTS, "Select n coplanar secondary Members", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 0, 8, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")

    ' controlled inputs
    Call pGeometricConstructionDefinition.AddControlledInput("AdvancedPlateSystem", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort1") ' relative to the first smart step
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1") ' relative to the first smart step
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort2") ' relative to the second smart step

    ' special controlled input to indicate that some PlateSystems of incoming Members have not been correctly un-bounded during the disconnect semantic,
    ' in the case, where the APS and some incoming Members are simultaneously deleted.
    ' this is due to the fact that FrameConnections depending from a deleted Member are first deleted, before being later recreated.
    ' our disconnect semantic executes, after those FrameConnections get deleted and before they get re-created.
    ' we have to wait the triggering of our compute semantic to have access again to the re-created ones.
    Call pGeometricConstructionDefinition.AddControlledInput("UnBoundedPlateSystems", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
     
    ' parameters
    Call pGeometricConstructionDefinition.AddParameter(sWELD_TOE, sWELD_TOE, GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter(sFILLET, sFILLET, GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter(sARM_LENGTH, sARM_LENGTH, GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter(sCUT_BACK, sCUT_BACK, GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter(sEXTENSION, sEXTENSION, GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
        
    ' code listed parameters
    Call pGeometricConstructionDefinition.AddParameter(sSUPPORT, "Side", GCCodeList, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue(sSUPPORT, "Back", 1)
    Call pGeometricConstructionDefinition.AddParameterValue(sSUPPORT, "Front", 2)
    
    ' parameter for cs orientation (added for mirror)
    Call pGeometricConstructionDefinition.AddParameter(sOrientation, "Orientation", GCCodeList, 0, 0, 0, 0, GCDirect)
    Call pGeometricConstructionDefinition.AddParameterValue(sOrientation, "Direct", GCDirect)
    Call pGeometricConstructionDefinition.AddParameterValue(sOrientation, "Indirect", GCIndirect)
    
    ' error codes
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MANDATORY_INPUT, "MissingMandatoryInput", "At least 2 members are required")
    Call pGeometricConstructionDefinition.AddErrorValue(COMMON_NODE_NOT_FOUND, "CommonNodeNotFound", "Common node not found")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXIS_TOO_SMALL, "MemberAxisTooSmall", "At least 1 member axis is too small")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXES_COLINEAR, "MemberAxesColinear", "Member axes colinear")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXES_NOT_COPLANAR, "MemberAxesNotCoplanar", "Member axes not coplanar")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_FACES_NOT_PARALLEL, "MemberFacesNotParallel", "Member faces not parallel")
    
    ' outputs
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Support")
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypePoint3d, "PseudoBoundary")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypePoint3d, "Node")
    Call pGeometricConstructionDefinition.AddOutput(GCLocalCoordinateSystem, "CoordinateSystem")
    
    Call DebugOut
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    
    Call ResetEntityCount
    On Error GoTo ErrorHandler
    
    ' use GC as a GCmacro
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGeometricConstruction
   
    ' retrieve info about the incoming members
    Dim pElementsOfPrimaryMembers As IJElements: Set pElementsOfPrimaryMembers = pGeometricConstruction.Inputs(sMEMBER_PARTS)
    Dim iCountOfPrimaryMembers As Integer: Let iCountOfPrimaryMembers = pElementsOfPrimaryMembers.Count
    Dim sKeysOfPrimaryMembers() As String: If iCountOfPrimaryMembers > 0 Then Let sKeysOfPrimaryMembers = Elements_GetKeys(pElementsOfPrimaryMembers)
    If iCountOfPrimaryMembers < 2 Then Err.Raise MISSING_MANDATORY_INPUT
            
    ' retrieve info about the secondary members
    Dim pElementsOfSecondaryMembers As IJElements: Set pElementsOfSecondaryMembers = pGeometricConstruction.Inputs(sSECONDARY_MEMBER_PARTS)
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
    
'''    ' retrieve lines of members axes
'''    Dim pLinesOfMemberAxes() As IJLine: Let pLinesOfMemberAxes = Members_GetLinesOfMemberSystemAxes(pElementsOfPrimaryMembers)
'''    If Lines_GetMinimumLength(pLinesOfMemberAxes) < EPSILON Then Err.Raise MEMBER_AXIS_TOO_SMALL
    
    ' compute common node
    Dim pPositionOfNode As IJDPosition: 'Set pPositionOfNode = GetPositionAtCommonExtremityOfLines(pLinesOfMemberAxes)
    Set pPositionOfNode = MemberParts_GetPositionOfCommonNode(pElementsOfPrimaryMembers)
    'If Not pPositionOfNode Is Nothing Then MsgBox (Position_ToString(pPositionOfNode, "CommonNode")) Else MsgBox "Nothing"
    If pPositionOfNode Is Nothing Then Err.Raise COMMON_NODE_NOT_FOUND
    
    ' retrieve lines of members axes
    Dim pLinesOfMemberAxes() As IJLine: Let pLinesOfMemberAxes = MemberParts_GetLinesOfMemberAxesAtCommonNode(pElementsOfPrimaryMembers, pPositionOfNode)
    If Lines_GetMinimumLength(pLinesOfMemberAxes) < EPSILON Then Err.Raise MEMBER_AXIS_TOO_SMALL
          
    ' compute support and boundaries - the array of keys for boundaries is passed ByRef and will be updated within the sub
    Call CreateNestedMacros(pPOM, pGeometricConstruction, _
                            pLinesOfMemberAxes, pPositionOfNode, _
                            iCountOfPrimaryMembers, pElementsOfPrimaryMembers, sKeysOfPrimaryMembers, _
                            iCountOfBoundaries, pElementsOfBoundaries, sKeysOfBoundaries)
    
    Call CreatePseudoBoundaries(pPOM, pGeometricConstruction, _
                                iCountOfSecondaryMembers, pElementsOfSecondaryMembers, sKeysOfSecondaryMembers, _
                                iCountOfPseudoBoundaries, pElementsOfPseudoBoundaries, sKeysOfPseudoBoundaries)
    
    ' if the APS is connected, then we can un-bound or bound added Members or modify the Plate to bound, if the context has changed (modify Support)
    If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        ' retrieve advanced plate system
        Dim oPlateSystemAdvanced As Object: Set oPlateSystemAdvanced = pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1)

        ' add new boundaries
        Call PlateSystem_AddNewBoundaries(oPlateSystemAdvanced, pGeometricConstructionMacro)
        
        ' get the key of the APS
        Dim sKey As String: Let sKey = pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").GetKey(oPlateSystemAdvanced)
       
        ' if the key is not set to "NoBound" then do the bound (this means that the APS has already computed its geometry and edge ports can be extracted)
        If Not sKey = "NotBounded" And Not sKey = "JustMigrated" Then
            ' bound/un-bound members
            'MsgBox "Bound"
            Call Members_BoundByOrUnboundFromEdgeOfAPS(pElementsOfPrimaryMembers, iCountOfPrimaryMembers, sKeysOfPrimaryMembers, _
                                                       pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                                       pGeometricConstruction, 1, oPlateSystemAdvanced, pPositionOfNode)
            Call Members_BoundByOrUnboundFromFaceOfAPS(pElementsOfSecondaryMembers, oPlateSystemAdvanced, pPositionOfNode)
        End If
        
        If sKey = "JustMigrated" Then
            Call Elements_RenameKeyOfElement(pGeometricConstruction.ControlledInputs("AdvancedPlateSystem"), "JustMigrated", "")
            ' force a recompute of the GCMacro
            Call pGeometricConstruction.UpdateGeometricConstructionAE
        End If
    End If
    
    ' remove unneeded boundaries
    If iCountOfBoundaries > 0 Then Call Elements_RemoveUnneeded(pElementsOfBoundaries, sKeysOfBoundaries)
    If iCountOfPseudoBoundaries > 0 Then Call Elements_RemoveUnneeded(pElementsOfPseudoBoundaries, sKeysOfPseudoBoundaries)
    
    Call DebugOut
    Exit Sub
ErrorHandler:
    Call GCProcessError(pGeometricConstruction, , Err.Number)
End Sub

Private Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal info As Variant, IsOnRibbonBar As Boolean)
    IsOnRibbonBar = True

    Select Case sName
        Case sOrientation: IsOnRibbonBar = False
    End Select
End Sub

'
' implementation of the IJSidePlate interface
'
Private Function IJSidePlate_Create(ByVal pPOM As IJDPOM, _
                                    ByVal pElementsOfMembers As IJElements, _
                                    ByVal vWeldToe As Variant, _
                                    ByVal vFillet As Variant, _
                                    ByVal vArmLength As Variant, _
                                    ByVal vCutBack As Variant, _
                                    ByVal vExtension As Variant, _
                                    Optional ByVal eSidePlateLocation As eSidePlateLocation = 1) As Object
    ' create factory
    Dim pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory
    Set pGeometricConstructionFactory = New GeometricConstructionEntitiesFactory

    ' create geometric construction
    Dim pGeometricConstruction As IJGeometricConstruction
    Set pGeometricConstruction = pGeometricConstructionFactory.CreateEntity("SidePlate", pPOM)
    
    ' provide inputs and parameters
    Dim i As Integer
    For i = 1 To pElementsOfMembers.Count
        pGeometricConstruction.Inputs(sMEMBER_PARTS).Add pElementsOfMembers(i), CStr(i)
    Next
    pGeometricConstruction.Parameter(sWELD_TOE) = vWeldToe
    pGeometricConstruction.Parameter(sFILLET) = vFillet
    pGeometricConstruction.Parameter(sARM_LENGTH) = vArmLength
    pGeometricConstruction.Parameter(sCUT_BACK) = vCutBack
    pGeometricConstruction.Parameter(sEXTENSION) = vExtension
    pGeometricConstruction.Parameter(sSUPPORT) = eSidePlateLocation
    
    ' return geometric construction
    Set IJSidePlate_Create = pGeometricConstruction
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
        On Error Resume Next ' avoid failure to not prevent the APS to get deleted
        Call Members_UnboundOnDelete(pGeometricConstruction, pRelationship.Destination)
        On Error GoTo 0

        ' delete the GCMacro itself, if no more processing is expected from the compute semantic
        If pGeometricConstruction.ControlledInputs("UnBoundedPlateSystems").Count = 0 Then Call Object_Delete(pGeometricConstruction)
    ElseIf Mid(pRelationship.Name, 1, Len(sMEMBER_PARTS)) = sMEMBER_PARTS And bIsDestinationDeleted Then
        ' an input member is deleting, disconnect the corresponding ports
        Dim sKey As String: Let sKey = Mid(pRelationship.Name, Len(sMEMBER_PARTS) + 2)
        pGeometricConstruction.ControlledInputs("AxisPort1").Remove (sKey)
        pGeometricConstruction.ControlledInputs("FacePort1").Remove (sKey)
    ElseIf Mid(pRelationship.Name, 1, Len(sSECONDARY_MEMBER_PARTS)) = sSECONDARY_MEMBER_PARTS And bIsDestinationDeleted Then
        ' an input member is deleting, disconnect the corresponding ports
        Let sKey = Mid(pRelationship.Name, Len(sSECONDARY_MEMBER_PARTS) + 2)
        pGeometricConstruction.ControlledInputs("AxisPort2").Remove (sKey)
    ElseIf Mid(pRelationship.Name, 1, Len(sMEMBER_PARTS)) = sMEMBER_PARTS Or Mid(pRelationship.Name, 1, Len(sSECONDARY_MEMBER_PARTS)) = sSECONDARY_MEMBER_PARTS Then
         ' Unbound removed member if the APS is connected
        If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
            Call Member_UnboundOnRemove(pGeometricConstruction, pRelationship.Destination)
        Else
            'input is removed clear controlled inputs if GC not deleted
            If bIsOriginDeleted = False Then
                pGeometricConstruction.ControlledInputs("AxisPort1").Clear
                pGeometricConstruction.ControlledInputs("FacePort1").Clear
                pGeometricConstruction.ControlledInputs("AxisPort2").Clear
            End If
        End If
        
    End If
    
    Call DebugOut
End Sub
'
' private APIs
'
Private Sub CreateNestedMacros(pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                       pLinesOfMemberAxes() As IJLine, pPositionOfNode As IJDPosition, _
                       iCountOfPrimaryMembers As Integer, pElementsOfPrimaryMembers As IJElements, sKeysOfPrimaryMembers() As String, _
                       iCountOfBoundaries As Integer, pElementsOfBoundaries As IJElements, ByRef sKeysOfBoundaries() As String)
    Call DebugIn(Method("CreateNestedMacros"))
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "CreateNestedMacros"
    
    Let sComment = "compute line normal to lines"
    Dim pLineOfNormal As IJLine: Set pLineOfNormal = GetLineAtNormalToLines(GetGCGeomFactory(), GetGCGeomFactory2(), pPOM, pLinesOfMemberAxes)
    If pLineOfNormal Is Nothing Then If iCountOfPrimaryMembers = 2 Then Err.Raise MEMBER_AXES_COLINEAR Else _
                                                                 Err.Raise MEMBER_AXES_NOT_COPLANAR
    
    Let sComment = "order lines"
    Dim lOrderedIndexes() As Long
    Dim dOrderedAngles() As Double
    Call Lines_OrderAroundPointAndNormal(pLinesOfMemberAxes, pPositionOfNode, Vector_FromLine(pLineOfNormal), lOrderedIndexes, dOrderedAngles)
        
    Let sComment = "create macros 'ExtractPorts'"
    Dim pGeometricConstructionMacrosOfExtractPorts() As IJGeometricConstructionMacro: ReDim pGeometricConstructionMacrosOfExtractPorts(1 To iCountOfPrimaryMembers)
    Dim iMember As Integer
    For iMember = 1 To iCountOfPrimaryMembers
        Set pGeometricConstructionMacrosOfExtractPorts(iMember) = CreateGeometricConstructionMacroOfExtractPorts(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                                 pElementsOfPrimaryMembers(iMember), sKeysOfPrimaryMembers(iMember), _
                                                                                                                 pLineOfNormal)
                                                                                                                         ' check if faces are coplanar
        If iMember > 1 Then If Not ArePlanesParallel(pGeometricConstructionMacrosOfExtractPorts(iMember).Outputs("Support").Item(1), _
                                                     pGeometricConstructionMacrosOfExtractPorts(1).Outputs("Support").Item(1)) Then _
            Err.Raise MEMBER_FACES_NOT_PARALLEL
    Next
    
    Let sComment = "populate output 'Node'"
    Call GeometricConstructionMacro_CreateNode(pGeometricConstruction, pPositionOfNode)
   
    Let sComment = "populate output 'Support'"
    Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Support", 1, pGeometricConstructionMacrosOfExtractPorts(1).Outputs("Support")(1))
    
    Let sComment = "populate output 'CoordinateSystem'"
    Call GeometricConstructionMacro_CreateCoordinateSystemOfSupport(pGeometricConstruction, pPOM, GetGCGeomFactory(), GetGCGeomFactory2(), pPositionOfNode)
    
    Let sComment = "create macros 'FreeEdge'"
    Dim pGeometricConstructionMacrosOfFreeEdges() As IJGeometricConstructionMacro: ReDim pGeometricConstructionMacrosOfFreeEdges(1 To iCountOfPrimaryMembers)
    Dim iFreeEdge As Integer
    For iFreeEdge = 1 To iCountOfPrimaryMembers
        ' index of the first Member supporting the FreeEdge
        Dim iMember1 As Integer: Let iMember1 = lOrderedIndexes(iFreeEdge)
        
        ' index of th second Member supporting the FreeEdge
        Dim iMember2 As Integer: If iFreeEdge < iCountOfPrimaryMembers Then Let iMember2 = lOrderedIndexes(iFreeEdge + 1) Else Let iMember2 = lOrderedIndexes(1)
                
        ' angle between the 2 Member axes
        Dim dAngle As Double: If iFreeEdge < iCountOfPrimaryMembers Then Let dAngle = dOrderedAngles(iFreeEdge + 1) - dOrderedAngles(iFreeEdge) Else Let dAngle = 360 - dOrderedAngles(iFreeEdge)

        ' compute "FreeEdge"
        Set pGeometricConstructionMacrosOfFreeEdges(iFreeEdge) = CreateGeometricConstructionMacroOfFreeEdge(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                            pElementsOfPrimaryMembers(iMember1), _
                                                                                                            pGeometricConstructionMacrosOfExtractPorts(iMember1), _
                                                                                                            pGeometricConstructionMacrosOfExtractPorts(iMember2), _
                                                                                                            dAngle)
        
        ' populate output "Boundary"
        Dim sKeyOfFreeEdge As String: Let sKeyOfFreeEdge = "FE1." + sKeysOfPrimaryMembers(iMember1) + "-" + sKeysOfPrimaryMembers(iMember2)
        Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Boundary", sKeyOfFreeEdge, pGeometricConstructionMacrosOfFreeEdges(iFreeEdge).Outputs("Boundary")(1))

        If iCountOfBoundaries > 0 Then
            Call Keys_RemoveKey(sKeysOfBoundaries, sKeyOfFreeEdge)
        End If
    Next
    
    Let sComment = "create macros 'TrimmingPoint' and 'TrimmingPlane'"
    For iMember = 1 To iCountOfPrimaryMembers
        ' index of the first FreeEdge used by the TrimmingPoint
        Dim iFreeEdge1 As Integer
        Dim j As Integer
        For j = 1 To iCountOfPrimaryMembers:
            If lOrderedIndexes(j) = iMember Then
                Let iFreeEdge1 = j:
                Exit For:
            End If
        Next
        
        ' index of the second FreeEdge used by the TrimmingPoint
        Dim iFreeEdge2 As Integer: If iFreeEdge1 = 1 Then Let iFreeEdge2 = iCountOfPrimaryMembers Else Let iFreeEdge2 = iFreeEdge1 - 1

        ' compute "TrimmingPoint"
        Dim pGeometricConstructionMacroOfTrimmingPoint As IJGeometricConstructionMacro
        Set pGeometricConstructionMacroOfTrimmingPoint = CreateGeometricConstructionMacroOfTrimmingPoint(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                         pGeometricConstructionMacrosOfExtractPorts(iMember), _
                                                                                                         pGeometricConstructionMacrosOfFreeEdges(iFreeEdge1), _
                                                                                                         pGeometricConstructionMacrosOfFreeEdges(iFreeEdge2))
        ' compute "TrimmingPlane"
        Dim pGeometricConstructionMacroOfTrimmingPlane As IJGeometricConstructionMacro
        Set pGeometricConstructionMacroOfTrimmingPlane = CreateGeometricConstructionMacroOfTrimmingPlane(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                         pElementsOfPrimaryMembers(iMember), _
                                                                                                         pGeometricConstructionMacroOfTrimmingPoint)
        ' populate output "Boundary"
        Dim sKeyOfTrimmingPlane As String: Let sKeyOfTrimmingPlane = "TR1." + sKeysOfPrimaryMembers(iMember)
        Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Boundary", sKeyOfTrimmingPlane, ModelBody_FromGeometry(pGeometricConstructionMacroOfTrimmingPlane.Outputs("Plane")(1)))
        
        If iCountOfBoundaries > 0 Then
            Call Keys_RemoveKey(sKeysOfBoundaries, sKeyOfTrimmingPlane)
        End If
    Next
    
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
'''Sub Members_UnboundOnDelete(pGeometricConstructionMacro As IJGeometricConstructionMacro, oAdvancedPlateSystem As Object)
'''    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Members_UnboundOnDelete"
'''
'''    ' use the GCMacro as a GC
'''    Dim pGeometricConstruction As IJGeometricConstruction: Set pGeometricConstruction = pGeometricConstructionMacro
'''
'''    Let sComment = "retrieve info about the primary members"
'''    Dim pElementsOfPrimaryMembers As IJElements: Set pElementsOfPrimaryMembers = Elements_GetValidElements(pGeometricConstruction.Inputs(sMEMBER_PARTS))
'''    Dim iCountOfPrimaryMembers As Integer: Let iCountOfPrimaryMembers = pElementsOfPrimaryMembers.Count
'''    Dim sKeysOfPrimaryMembers() As String: If iCountOfPrimaryMembers > 0 Then Let sKeysOfPrimaryMembers = Elements_GetDummyKeys(pElementsOfPrimaryMembers)
'''
'''    Let sComment = "retrieve position of node"
'''    Dim pPositionOfNode As IJDPosition
'''    If iCountOfPrimaryMembers > 0 Then
'''        ' if something goes wrong, we will at least try to retrieve the node
'''        On Error Resume Next
'''
'''        ' retrieve lines of member axes
'''        Dim pLinesOfMemberAxes() As IJLine: Let pLinesOfMemberAxes = Members_GetLinesOfMemberSystemAxes(pElementsOfPrimaryMembers)
'''
'''        ' compute the node
'''        Set pPositionOfNode = GetPositionAtCommonExtremityOfLines(pLinesOfMemberAxes)
'''
'''        On Error GoTo 0
'''
'''        ' retrieve the node, if not computed
'''        If pPositionOfNode Is Nothing Then Set pPositionOfNode = GeometricConstructionMacro_RetrievePositionOfNode(pGeometricConstruction)
'''    End If
'''
'''    If iCountOfPrimaryMembers > 0 Then
'''        Let sComment = "retrieve info about the boundaries"
'''        Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
'''        Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
'''        Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)
'''
'''        ' un-bound removed primary members
'''        Call Members_UnboundFromEdgeOfAPS(pElementsOfPrimaryMembers, iCountOfPrimaryMembers, sKeysOfPrimaryMembers, _
'''                                          pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
'''                                          pGeometricConstruction, 1, oAdvancedPlateSystem, pPositionOfNode)
'''    End If
'''
'''    Let sComment = "retrieve info about the secondary members"
'''    Dim pElementsOfSecondaryMembers As IJElements: Set pElementsOfSecondaryMembers = Elements_GetValidElements(pGeometricConstruction.Inputs(sSECONDARY_MEMBER_PARTS))
'''    Dim iCountOfSecondaryMembers As Integer: Let iCountOfSecondaryMembers = pElementsOfSecondaryMembers.Count
'''    Dim sKeysOfSecondaryMembers() As String: If iCountOfSecondaryMembers > 0 Then Let sKeysOfSecondaryMembers = Elements_GetDummyKeys(pElementsOfSecondaryMembers)
'''
'''    If iCountOfSecondaryMembers > 0 Then
'''        Let sComment = "retrieve info about the boundaries"
'''        Dim pElementsOfPseudoBoundaries As IJElements: Set pElementsOfPseudoBoundaries = pGeometricConstructionMacro.Outputs("PseudoBoundary")
'''        Dim iCountOfPseudoBoundaries As Integer: Let iCountOfPseudoBoundaries = pElementsOfPseudoBoundaries.Count
'''        Dim sKeysOfPseudoBoundaries() As String: If iCountOfPseudoBoundaries > 0 Then Let sKeysOfPseudoBoundaries = Elements_GetKeys(pElementsOfPseudoBoundaries)
'''
'''        Let sComment = "un-bound removed secondary members"
'''        Call Members_UnboundFromFaceOfAPS(pElementsOfSecondaryMembers, iCountOfSecondaryMembers, sKeysOfSecondaryMembers, _
'''                                          pElementsOfPseudoBoundaries, iCountOfPseudoBoundaries, sKeysOfPseudoBoundaries, _
'''                                          pGeometricConstruction, 2, oAdvancedPlateSystem, pPositionOfNode)
'''
'''    End If
'''    Exit Sub
'''ErrorHandler:
'''    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
'''    Call ShowError(sContext, sComment, Err.Number, Err.Description)
'''    Err.Raise Err.Number, sContext, sComment
'''End Sub
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
        'If Mid(sName, 1, Len("MemberPart")) = "MemberPart" Then Call GeometricConstruction_MigrateInputs(MyGC, pMigrateHelper, sName, pPositionOfNode)
        If InStr(1, sName, "MemberPart", vbTextCompare) > 0 Then Call GeometricConstruction_MigrateInputs(MyGC, pMigrateHelper, sName, pPositionOfNode)
    Next

    Call DebugOut
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
 '   Adaptation of GCSidePlate GC:
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


