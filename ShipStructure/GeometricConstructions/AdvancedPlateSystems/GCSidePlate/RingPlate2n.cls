VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RingPlate2n"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_oGenericPlateHelper As Object

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

' record version number
Dim m_bAreV2011Metadata As Boolean

' define factories
Dim m_pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory

Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGCMigrate
Implements IJGeometricConstructionDynamic_IsOnRibbonBar
Implements IJGCToDoDelegate
Implements IJGCConnectedElementSC

Dim m_oAdvancedDebug As New AdvancedDebug
Private Property Get Source() As String
    Let Source = "GCAPSWithTrimBack3.GenericPlateImp"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property
'
' constructor/destructor of the class
'
Private Sub Class_Initialize()
    Call m_oAdvancedDebug.StartSource(Source)
    Set m_pGeometricConstructionFactory = New GeometricConstructionEntitiesFactory

    Set m_oGenericPlateHelper = CreateObject("GCGenericPlateHelper.GenericPlateHelper")
End Sub
Private Sub Class_Terminate()
    Set m_oGenericPlateHelper = Nothing
    Set m_pGeometricConstructionFactory = Nothing
    
    Call m_oAdvancedDebug.StopSource(Source)
    Set m_oAdvancedDebug = Nothing
End Sub

Private Sub IJGCConnectedElementSC_PostDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCConnectedElementSC_PostDisconnectExternalRels"))
    
    Call m_oGenericPlateHelper.PostDisconnectExternalRels(pGC, pInfo)
    
    Call m_oAdvancedDebug.ExitMethod

End Sub

Private Sub IJGCConnectedElementSC_PreDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCConnectedElementSC_PreDisconnectExternalRels"))
    
    Call m_oGenericPlateHelper.PreDisconnectExternalRels(pGC, pInfo)
    
    Call m_oAdvancedDebug.ExitMethod

End Sub

'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGCDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Initialize"))
    
    Call m_oGenericPlateHelper.PreInitialize(pGCDefinition)
    Call Initialize(pGCDefinition)
    Call m_oGenericPlateHelper.PostInitialize(pGCDefinition)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGC As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    
    Dim bEvaluateGeometry As Boolean
    Call m_oGenericPlateHelper.PreEvaluate(pGC, pPOM, bEvaluateGeometry)
    Call m_oAdvancedDebug.EnterMethod(Method("EvaluateGeometry"))
    If bEvaluateGeometry Then
        Call Evaluate(pGC, pPOM)
    Else
        Call m_oAdvancedDebug.ShowMsg("Skip")
    End If
    Call m_oAdvancedDebug.ExitMethod
    Call m_oGenericPlateHelper.PostEvaluate(pGC, pPOM)

    Call m_oAdvancedDebug.ExitMethod
    Exit Sub
ErrorHandler:
    Call m_oAdvancedDebug.ExitMethod
    Call m_oGenericPlateHelper.ProcessError(pGC, Err.Number)
End Sub
'
' implementation of the IJGCSemanticConnection interface
'
Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal oRelationship As Object)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PostConnectionAdded"))
    
    Call m_oGenericPlateHelper.PostConnectionAdded(oRelationship)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PreConnectionRemoved"))
    
    Call m_oGenericPlateHelper.PreConnectionRemoved(oRelationship, bIsOriginDeleted, bIsDestinationDeleted)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCMigrate interface
'
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCMigrate_Migrate"))

    Call m_oGenericPlateHelper.Migrate(MyGC, pMigrateHelper)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGeometricConstructionDynamic_IsOnRibbonBar interface
'
Public Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As GeometricConstruction, ByVal vInfo As Variant, ByRef bIsOnRibbonBar As Boolean)
    Call m_oGenericPlateHelper.IsOnRibbonBar(sName, pGC, vInfo, bIsOnRibbonBar)
End Sub
'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Initialize"))
    
    ' located inputs
    Call pGeometricConstructionDefinition.AddInput("MemberParts", "Select n coplanar Members or edge reinforcements in consecutive counterclockwise or clockwise order", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 100, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify IJSplitNotify IJStructSplit IJDProfileMoldedConventions IJStiffenerAttachmentMethod ISPSMemberSystemXSectionNotify IJSymbolUpdateInt ISPSMemberSystemSuppingNotify1")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfMemberParts", "BehaviorOfMemberParts", 1, 0, 0, 0, 0, "MembersBoundByEdgesOfPlate", 0)
    Call pGeometricConstructionDefinition.AddInput("OrthogonalMemberPart", "Select 1 orthogonal Member", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfOrthogonalMemberPart", "BehaviorOfOrthogonalMemberPart", 1, 0, 0, 0, 0, "MembersBoundPlate,MembersBoundByPlate,MembersPenetrateByPlate", 0)
    Call pGeometricConstructionDefinition.AddInput("SecondaryMemberParts", "Select n coplanar secondary Members", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 0, 8, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")

    ' controlled inputs
    'Call pGeometricConstructionDefinition.AddControlledInput("AdvancedPlateSystem", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort2")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort2")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort3") ' relative to the third smart step

    ' special controlled input to indicate that some PlateSystems of incoming Members have not been correctly un-bounded during the disconnect semantic,
    ' in the case, where the APS and some incoming Members are simultaneously deleted.
    ' this is due to the fact that FrameConnections depending from a deleted Member are first deleted, before being later recreated.
    ' our disconnect semantic executes, after those FrameConnections get deleted and before they get re-created.
    ' we have to wait the triggering of our compute semantic to have access again to the re-created ones.
    Call pGeometricConstructionDefinition.AddControlledInput("UnBoundedPlateSystems", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
     ' code listed parameters
    Call pGeometricConstructionDefinition.AddParameter("Support", "Support", GCCodeList, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Back", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Front", 2)
   
   ' parameters
    Call pGeometricConstructionDefinition.AddParameter("WeldToe", "WeldToe", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1.5)
    Call pGeometricConstructionDefinition.AddParameter("Fillet", "Fillet", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("CutBack", "CutBack", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("CutBack2", "CutBack2", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("SlopeRatio", "SlopeRatio", 8, UNIT_SCALAR, 0, 0, 0, 1.01)
    Call pGeometricConstructionDefinition.AddParameter("Extension", "Extension", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    
    ' code listed parameters
    Call pGeometricConstructionDefinition.AddParameter("Support", "Support", GCCodeList, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Back", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Front", 2)
    
'''    If m_bAreV2011Metadata Then
'''        Call pGeometricConstructionDefinition.AddParameter("IsCutByTube", "IsCutByTube", GCCodeList, 0, 0, 0, 0, 1, True)
'''        Call pGeometricConstructionDefinition.AddParameterValue("IsCutByTube", "False", 0)
'''        Call pGeometricConstructionDefinition.AddParameterValue("IsCutByTube", "True", 1)
'''    End If
    
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
    
    Call DebugOut
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    
    Call ResetEntityCount
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
    
    ' compute or retrieve common node
    Dim pLineOfOrthogonalMemberAxis As IJLine: Set pLineOfOrthogonalMemberAxis = MemberPart_GetLine(oOrthogonalMember)
    Dim pPositionOfNode As IJDPosition: Set pPositionOfNode = MemberParts_GetPositionOfCommonNodeOnCurve(pElementsOfMembers, pLineOfOrthogonalMemberAxis)
    If pPositionOfNode Is Nothing Then Err.Raise COMMON_NODE_NOT_FOUND
    
    ' retrieve lines of members axes
    Dim pLinesOfMemberAxes() As IJLine: Let pLinesOfMemberAxes = MemberParts_GetLinesOfMemberAxesAtCommonNode(pElementsOfMembers, pPositionOfNode)
    If Lines_GetMinimumLength(pLinesOfMemberAxes) < EPSILON _
    Or pLineOfOrthogonalMemberAxis.Length() < EPSILON Then Err.Raise MEMBER_AXIS_TOO_SMALL
    
    ' compute support and boundaries - the array of keys for boundaries is passed ByRef and will be updated within the sub
    Call CreateNestedMacros(pPOM, pGeometricConstruction, _
                            pLinesOfMemberAxes, pPositionOfNode, _
                            pLineOfOrthogonalMemberAxis, _
                            iCountOfMembers, pElementsOfMembers, sKeysOfMembers, _
                            iCountOfBoundaries, pElementsOfBoundaries, sKeysOfBoundaries)
            
    Call CreatePseudoBoundaries(pPOM, pGeometricConstruction, _
                                iCountOfSecondaryMembers, pElementsOfSecondaryMembers, sKeysOfSecondaryMembers, _
                                iCountOfPseudoBoundaries, pElementsOfPseudoBoundaries, sKeysOfPseudoBoundaries)
    
   
    ' remove unneeded boundaries
    If iCountOfBoundaries > 0 Then Call Elements_RemoveUnneeded(pElementsOfBoundaries, sKeysOfBoundaries)
    If iCountOfPseudoBoundaries > 0 Then Call Elements_RemoveUnneeded(pElementsOfPseudoBoundaries, sKeysOfPseudoBoundaries)
    
    Call DebugOut
    Exit Sub
ErrorHandler:
    Call GCProcessError(pGeometricConstruction, , Err.Number)
End Sub
''''
'''' implementation of the IJGCSemanticConnection interface
''''
'''Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal oRelationship As Object)
'''    Call DebugIn(Method("IJGCSemanticConnection_PostConnectionAdded"))
'''
'''    Dim pRelationship As IJDRelationship: Set pRelationship = oRelationship
'''    Call DebugInput("Origin", pRelationship.Origin)
'''    Call DebugInput("Destination", pRelationship.Destination)
'''    Call DebugInput("Name", pRelationship.Name)
'''
'''    Call DebugOut
'''End Sub
'''Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
'''    Call DebugIn(Method("IJGCSemanticConnection_PreConnectionRemoved"))
'''
'''    Dim pRelationship As IJDRelationship: Set pRelationship = oRelationship
'''    Call DebugInput("Origin", pRelationship.Origin)
'''    Call DebugInput("Destination", pRelationship.Destination)
'''    Call DebugInput("Name", pRelationship.Name)
'''    Call DebugInput("IsOriginDeleted", bIsOriginDeleted)
'''    Call DebugInput("IsDestinationDeleted", bIsDestinationDeleted)
'''
'''    Dim pGeometricConstruction As IJGeometricConstruction: Set pGeometricConstruction = pRelationship.Origin
'''
'''    If Mid(pRelationship.Name, 1, Len("AdvancedPlateSystem")) = "AdvancedPlateSystem" And bIsDestinationDeleted Then
'''        ' the APS is deleting
'''        ' unbound remaining connected Members
'''        On Error Resume Next ' avoid failure to not prevent the APS to get deleted
'''        Call Members_UnboundOnDelete(pGeometricConstruction, pRelationship.Destination)
'''        On Error GoTo 0
'''
'''        ' delete the GCMacro itself, if no more processing is expected from the compute semantic
'''        If pGeometricConstruction.ControlledInputs("UnBoundedPlateSystems").Count = 0 Then Call Object_Delete(pGeometricConstruction)
'''    ElseIf Mid(pRelationship.Name, 1, Len("MemberParts")) = "MemberParts" And bIsDestinationDeleted Then
'''        ' an input member is deleting, disconnect the corresponding ports
'''        Dim sKey As String: Let sKey = Mid(pRelationship.Name, Len("MemberParts") + 2)
'''        pGeometricConstruction.ControlledInputs("AxisPort1").Remove (sKey)
'''        pGeometricConstruction.ControlledInputs("FacePort1").Remove (sKey)
'''    End If
'''
'''    Call DebugOut
'''End Sub
'
' private APIs
'
'''Private Sub UnboundMembersOnDelete(pGeometricConstructionMacro As IJGeometricConstructionMacro, oAdvancedPlateSystem As Object)
'''    ' use the GCMacro as a GC
'''    Dim pGeometricConstruction As IJGeometricConstruction: Set pGeometricConstruction = pGeometricConstructionMacro
'''
'''    ' retrieve info about the incoming members
'''    Dim pElementsOfMembers As IJElements: Set pElementsOfMembers = Elements_GetValidElements(pGeometricConstruction.Inputs("MemberParts"))
'''    Dim iCountOfMembers As Integer: Let iCountOfMembers = pElementsOfMembers.Count
'''    Dim sKeysOfMembers() As String: If iCountOfMembers > 0 Then Let sKeysOfMembers = Elements_GetDummyKeys(pElementsOfMembers)
'''
'''    ' retrieve info about the orthogonal member
'''    Dim oOrthogonalMember As Object
'''    If True Then
'''        Dim pElementsOfValidElements As IJElements:
'''        Set pElementsOfValidElements = Elements_GetValidElements(pGeometricConstruction.Inputs("OrthogonalMemberPart"))
'''        If pElementsOfValidElements.Count = 1 Then Set oOrthogonalMember = pElementsOfValidElements(1)
'''    End If
'''
'''    If iCountOfMembers > 0 Then
'''        ' retrieve info about the boundaries
'''        Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
'''        Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
'''        Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)
'''
'''        ' compute the node
'''        Dim pPositionOfNode As IJDPosition
'''        If Not oOrthogonalMember Is Nothing Then
'''            ' if something goes wrong, we will at least try to retrieve the node
'''            On Error Resume Next
'''
'''            ' retrieve axes of incoming members
'''            Dim pLinesOfMemberAxes() As IJLine: Let pLinesOfMemberAxes = Members_GetLinesOfMemberSystemAxes(pElementsOfMembers)
'''
'''            ' retrieve axis of orthogonal member
'''            Dim pLineOfOrthogonalMemberAxis As IJLine
'''            If Not oOrthogonalMember Is Nothing Then Set pLineOfOrthogonalMemberAxis = MemberPart_GetLine(oOrthogonalMember)
'''
'''            If Lines_GetMinimumLength(pLinesOfMemberAxes) > EPSILON _
'''            And pLineOfOrthogonalMemberAxis.Length() > EPSILON Then
'''                ' compute common node
'''                Set pPositionOfNode = ComputePositionOfNode(pGeometricConstructionMacro, pLinesOfMemberAxes, iCountOfMembers, pLineOfOrthogonalMemberAxis)
'''            End If
'''
'''            On Error GoTo 0
'''        End If
'''
'''        ' retrieve the node, if not computed
'''        If pPositionOfNode Is Nothing Then Set pPositionOfNode = GeometricConstructionMacro_RetrievePositionOfNode(pGeometricConstruction)
'''
'''        ' un-bound removed members
'''        Call UnboundMembersAtNode(pGeometricConstruction, 1, _
'''                                  iCountOfMembers, pElementsOfMembers, sKeysOfMembers, _
'''                                  iCountOfBoundaries, pElementsOfBoundaries, sKeysOfBoundaries, _
'''                                  oAdvancedPlateSystem, pPositionOfNode)
'''    End If
'''End Sub
Private Function ComputePositionOfNode(pGeometricConstructionMacro As IJGeometricConstructionMacro, _
                                       pLinesOfMemberAxes() As IJLine, iCountOfMembers As Integer, _
                                       pLineOfOrthogonalMemberAxis As IJLine) As IJDPosition
    Call DebugIn(Method("ComputePositionOfNode"))
    
    ' prepare result
    Dim pPositionOfNode As IJDPosition
    
    ' compute common point on curve
    Set pPositionOfNode = GetPositionAtExtremityOfLineOnCurve(pLinesOfMemberAxes(1), pLineOfOrthogonalMemberAxis)
    
    ' verify that common node also works for regular members
    If Not pPositionOfNode Is Nothing And iCountOfMembers > 1 Then
        Dim pPositionOfNode1 As IJDPosition: Set pPositionOfNode1 = GetPositionAtCommonExtremityOfLines(pLinesOfMemberAxes)
        If Not pPositionOfNode1 Is Nothing Then If pPositionOfNode1.DistPt(pPositionOfNode) > EPSILON Then Set pPositionOfNode = Nothing
    End If
    
    ' return result
    Set ComputePositionOfNode = pPositionOfNode
    
    Call DebugOut
End Function
Private Function IsCardinalPointOfMembersValid(pElementsOfMembers As IJElements, iCountOfMembers As Integer) As Boolean
    Call DebugIn(Method("IsCardinalPointOfMembersValid"))
    
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

    Call DebugOut
End Function
Private Sub CreateNestedMacros(pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                       pLinesOfMemberAxes() As IJLine, pPositionOfNode As IJDPosition, _
                       pLineOfOrthogonalLine As IJLine, _
                       iCountOfMembers As Integer, pElementsOfMembers As IJElements, sKeysOfMembers() As String, _
                       iCountOfBoundaries As Integer, pElementsOfBoundaries As IJElements, ByRef sKeysOfBoundaries() As String)
    Call DebugIn(Method("CreateNestedMacros"))
    
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
    
'''    ' populate first output "Boundary"
'''    Dim bIsCutByTube As Boolean: Let bIsCutByTube = True
'''    If m_bAreV2011Metadata Then bIsCutByTube = CInt(pGeometricConstruction.Parameter("IsCutByTube")) = 1
'''    If bIsCutByTube Then
'''        Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Boundary", "FacePort2", pGeometricConstructionMacroOfExtractCircle.Outputs("Boundary")(1))
'''    Else
'''        Call GeometricConstructionMacro_DeleteOutput(pGeometricConstruction, "Boundary", "FacePort2")
'''    End If
    
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
        Dim sCutBack1 As String
        Dim sCutBack2 As String
        If iMember1 Mod 2 = 1 Then
            ' odd member
            Let sCutBack1 = "CutBack"
            Let sCutBack2 = "CutBack2"
        Else
            ' even member
            Let sCutBack1 = "CutBack2"
            Let sCutBack2 = "CutBack"
        End If
        Set pGeometricConstructionMacrosOfFreeEdges(iFreeEdge) = CreateGeometricConstructionMacroOfStraightEdgeAround2(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                                      pElementsOfMembers(iMember1), sCutBack1, sCutBack2, _
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

        ' compute "TrimmingPoint"
        Dim pGeometricConstructionMacroOfTrimmingPoint As IJGeometricConstructionMacro
        If pGeometricConstruction.Parameter("SlopeRatio") <> 0 Then
            Set pGeometricConstructionMacroOfTrimmingPoint = CreateGeometricConstructionMacroOfTrimmingPoint2(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                         pGeometricConstructionMacrosOfExtractPorts(iMember), _
                                                                                                         pGeometricConstructionMacrosOfFreeEdges(iFreeEdge1), _
                                                                                                         pGeometricConstructionMacrosOfFreeEdges(iFreeEdge2))
        Else
            Set pGeometricConstructionMacroOfTrimmingPoint = CreateGeometricConstructionMacroOfTrimmingPoint(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                         pGeometricConstructionMacrosOfExtractPorts(iMember), _
                                                                                                         pGeometricConstructionMacrosOfFreeEdges(iFreeEdge1), _
                                                                                                         pGeometricConstructionMacrosOfFreeEdges(iFreeEdge2))
        End If
        
        ' compute "TrimmingPlane"
        Dim pGeometricConstructionMacroOfTrimmingPlane As IJGeometricConstructionMacro
        Set pGeometricConstructionMacroOfTrimmingPlane = CreateGeometricConstructionMacroOfTrimmingPlane(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                         pElementsOfMembers(iMember), _
                                                                                                         pGeometricConstructionMacroOfTrimmingPoint)
        
        ' populate output "Boundary" for trimming plane
        Dim sKeyOfTrimmingPlane As String: Let sKeyOfTrimmingPlane = "TR1." + sKeysOfMembers(iMember)
        Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Boundary", sKeyOfTrimmingPlane, ModelBody_FromGeometry(pGeometricConstructionMacroOfTrimmingPlane.Outputs("Plane")(1)))
        If iCountOfBoundaries > 0 Then Call Keys_RemoveKey(sKeysOfBoundaries, sKeyOfTrimmingPlane)
        
        ' compute "SnipedSurfaces"
        If pGeometricConstruction.Parameter("SlopeRatio") <> 0 Then
            Dim pGeometricConstructionMacroOfSnipedSurfaces As IJGeometricConstructionMacro
            Set pGeometricConstructionMacroOfSnipedSurfaces = CreateGeometricConstructionMacroOfSnipedSurfaces(m_pGeometricConstructionFactory, pPOM, pGeometricConstruction, _
                                                                                                             pElementsOfMembers(iMember), _
                                                                                                             pGeometricConstructionMacrosOfExtractPorts(iMember), _
                                                                                                             pGeometricConstructionMacroOfTrimmingPoint)
        
            ' populate output "Boundary" for sniped surface left
            Dim sKeyOfSnipedSurfaceLeft As String: Let sKeyOfSnipedSurfaceLeft = "SL1." + sKeysOfMembers(iMember)
            Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Boundary", sKeyOfSnipedSurfaceLeft, pGeometricConstructionMacroOfSnipedSurfaces.Outputs("SurfaceLeft")(1))
            If iCountOfBoundaries > 0 Then Call Keys_RemoveKey(sKeysOfBoundaries, sKeyOfSnipedSurfaceLeft)
                        
            ' populate output "Boundary" for sniped surface right
            Dim sKeyOfSnipedSurfaceRight As String: Let sKeyOfSnipedSurfaceRight = "SR1." + sKeysOfMembers(iMember)
            Call GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstruction, "Boundary", sKeyOfSnipedSurfaceRight, pGeometricConstructionMacroOfSnipedSurfaces.Outputs("SurfaceRight")(1))
            If iCountOfBoundaries > 0 Then Call Keys_RemoveKey(sKeysOfBoundaries, sKeyOfSnipedSurfaceRight)
        End If
    Next
    
    Call DebugOut
End Sub
''''
'''' implementation of the IJGCMigrate interface
''''
'''Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
'''    Call DebugIn(Method("IJGCMigrate_Migrate"))
'''
'''    ' retrieve position of node
'''    Dim pPositionOfNode As IJDPosition
'''    If True Then
'''        Dim MyGCMacro As IJGeometricConstructionMacro: Set MyGCMacro = MyGC
'''        Set pPositionOfNode = Position_FromPoint(MyGCMacro.Outputs("Node")(1))
'''    End If
'''
'''    ' loop on all the inputs, whose name is prefixed by "MemberPart"
'''    Dim pGeometricConstructionDefinition As IJGeometricConstructionDefinition: Set pGeometricConstructionDefinition = MyGC.definition
'''    Dim i  As Integer
'''    For i = 1 To pGeometricConstructionDefinition.InputCount
'''        Dim sName As String, sPrompt As String, sFilter As String, lMinConnected As Long, lMaxConnected As Long, sComputeIIDs As String
'''        Call pGeometricConstructionDefinition.GetInputInfoByIndex(i, sName, sPrompt, sFilter, lMinConnected, lMaxConnected, sComputeIIDs)
'''
'''        ' migrate these inputs, if one of them has been replaced
'''        'If Mid(sName, 1, Len("MemberPart")) = "MemberPart" Then Call GeometricConstruction_MigrateInputs(MyGC, pMigrateHelper, sName, pPositionOfNode)
'''        If InStr(1, sName, "MemberPart", vbTextCompare) > 0 Then Call GeometricConstruction_MigrateInputs(MyGC, pMigrateHelper, sName, pPositionOfNode)
'''    Next
'''
'''    Call DebugOut
'''End Sub
'''
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCToDoDelegate_ToDoDelegate"))

    Set IJGCToDoDelegate_ToDoDelegate = Nothing
    On Error Resume Next
    Set IJGCToDoDelegate_ToDoDelegate = m_oGenericPlateHelper.ToDoDelegate(pGC)
    On Error GoTo 0

    Call m_oAdvancedDebug.ExitMethod
End Property





