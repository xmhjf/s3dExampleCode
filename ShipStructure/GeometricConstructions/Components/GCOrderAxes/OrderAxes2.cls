VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderAxes2"
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
    MEMBER_AXES_NOT_ORTHOGONAL = 7
    COMMON_NODE_NOT_SHARED = 8
End Enum

Const m_sModule = "GCOrderAxes.OrderAxes2"
Const sPRIMARY_PROFILES = "PrimaryProfiles"
Const sSECONDARY_PROFILES = "SecondaryProfiles"
Const sCOORDINATE_SYSTEM = "CoordinateSystem"

Implements IJGeometricConstructionDefinitionService
Private Property Get Source() As String
    Let Source = "GCOrderAxes.OrderAxes2"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property
Private Sub Class_Initialize()
    Call DebugStart(Source)
End Sub
Private Sub Class_Terminate()
    Call DebugStop(Source)
End Sub
'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Initialize: OrderAxes2"))
    
    ' located inputs
    Call pGeometricConstructionDefinition.AddInput(sPRIMARY_PROFILES, "Select n coplanar primary Members, defining the support plane of the APS", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem] AND [GCFilters.Filters,IsLinearStiffenerSystem]", 1, 8, "")
    Call pGeometricConstructionDefinition.AddInput(sSECONDARY_PROFILES, "Select n secondary Members, defining the normal vector of the APS", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 0, 8, "")
    Call pGeometricConstructionDefinition.AddInput(sCOORDINATE_SYSTEM, "Select an optional CS defining the orientation of the normal", sFILTER_IS_COORDINATE_SYSTEM, 0, 1)
        
    ' controlled inputs
    Call pGeometricConstructionDefinition.AddControlledInput("Port1", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port2", "ISPSSplitAxisAlongPort")
    
    ' code listed parameters
    Call pGeometricConstructionDefinition.AddParameter("Support", "Direction", GCCodeList, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Negative", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Positive", 2)
    
    ' error codes
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MANDATORY_INPUT, "MissingMandatoryInput", "At least 2 members are required")
    Call pGeometricConstructionDefinition.AddErrorValue(COMMON_NODE_NOT_FOUND, "CommonNodeNotFound", "Common node not found")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXIS_TOO_SMALL, "MemberAxisTooSmall", "At least 1 member axis is too small")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXES_COLINEAR, "MemberAxesColinear", "Member axes colinear")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXES_NOT_COPLANAR, "MemberAxesNotCoplanar", "Member axes not coplanar")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_FACES_NOT_PARALLEL, "MemberFacesNotParallel", "Member faces not parallel")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXES_NOT_ORTHOGONAL, "MemberAxesNotOrthogonal", "Member axes not perpendicular to the orthogonal member")
    Call pGeometricConstructionDefinition.AddErrorValue(COMMON_NODE_NOT_SHARED, "CommonNodeNotShared", "Common node not shared by orthogonal or column members")
    
    ' outputs
    Call pGeometricConstructionDefinition.AddOutput(GCGTypeLine3d, "Axis")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypeLine3d, "Normal")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypePoint3d, "Node")
    Call pGeometricConstructionDefinition.AddOutput(GCLocalCoordinateSystem, "CoordinateSystem")
    
    Call DebugOut
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    On Error GoTo ErrorHandler
    
    Dim pElementsOfPrimaryProfiles As IJElements: Set pElementsOfPrimaryProfiles = pGeometricConstruction.Inputs(sPRIMARY_PROFILES)
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGeometricConstruction
'''    If TypeOf pElementsOfPrimaryProfiles(1) Is ISPSMemberSystemLinear _
'''    And TypeOf pElementsOfPrimaryProfiles(2) Is ISPSMemberSystemLinear Then
'''        Dim oCSByLines0 As Object
'''        Set oCSByLines0 = GetGCGeomFactory().CSByLines.PlaceGeometry(Nothing, pElementsOfPrimaryProfiles(1), pElementsOfPrimaryProfiles(2), GCXY, GCDirect)
'''        Dim oPointByCurves0 As Object
'''        Set oPointByCurves0 = GetGCGeomFactory().PointByCurves.PlaceGeometry(Nothing, pElementsOfPrimaryProfiles(1), pElementsOfPrimaryProfiles(2), Nothing, Nothing)
'''
'''        pGeometricConstructionMacro.Output("CoordinateSystem", 1) = oCSByLines0
'''        pGeometricConstructionMacro.Output("Node", 1) = oPointByCurves0
'''
'''        Exit Sub
'''    End If
'''
'''    ' use GC as a GCmacro
    
    ' retrieve incoming Members
    Dim pElementsOfSecondaryProfiles As IJElements: Set pElementsOfSecondaryProfiles = pGeometricConstruction.Inputs(sSECONDARY_PROFILES)
    If pElementsOfPrimaryProfiles.Count = 0 Then Err.Raise MISSING_MANDATORY_INPUT

    ' extract LineAxisPort of each incoming Member
    Dim pGCMemberFactory As IJGCMemberFactory: Set pGCMemberFactory = CreateObject("GCSPSDefinitions.GCMemberFactory")
    If True Then
        Dim pGeometricConstructionEntitiesFactory As IJGeometricConstructionEntitiesFactory: Set pGeometricConstructionEntitiesFactory = New GeometricConstructionEntitiesFactory
        pGeometricConstruction.ControlledInputs("Port1").Clear
        Dim i As Integer
        For i = 1 To pElementsOfPrimaryProfiles.Count
            Dim oProfile As Object: Set oProfile = pElementsOfPrimaryProfiles(i)
            Dim sKey As String: Let sKey = pElementsOfPrimaryProfiles.GetKey(oProfile)
            
            Dim pGCEntitiesFactory As IJGeometricConstructionEntitiesFactory: Set pGCEntitiesFactory = New GeometricConstructionEntitiesFactory
            
            If TypeOf oProfile Is IJStiffenerSystem Then
                Dim oAxisPortExtractor As IJGeometricConstruction
                Set oAxisPortExtractor = pGCEntitiesFactory.CreateEntity("AxisPortExtractor", pPOM, "00" + CStr(i) + "-AxisPortExtractor")
                oAxisPortExtractor.Input("Connectable") = oProfile
                oAxisPortExtractor.Parameter("GeometrySelector") = 4 ' GCStable
                oAxisPortExtractor.Evaluate
                
                pGeometricConstruction.ControlledInputs("Port1").Add oAxisPortExtractor.ControlledInputs("Port")(1), sKey
           ElseIf Not TypeOf oProfile Is ISPSMemberSystemLinear Then
                Dim oLineAxisPortExtractor As IJGeometricConstruction
                Set oLineAxisPortExtractor = pGCEntitiesFactory.CreateEntity("LineAxisPortExtractor", pPOM, "00" + CStr(i) + "-LineAxisPortExtractor")
                oLineAxisPortExtractor.Input("MemberPart") = oProfile
                Call DebugMsg("Extract port from MemberPart")
                oLineAxisPortExtractor.Evaluate
                
                pGeometricConstruction.ControlledInputs("Port1").Add oLineAxisPortExtractor.ControlledInputs("Port")(1), sKey
            Else
                ' ISPSMemberSystemLinear : add the member itself
                pGeometricConstruction.ControlledInputs("Port1").Add oProfile, sKey
            End If
        Next
        pGeometricConstruction.ControlledInputs("Port2").Clear
        For i = 1 To pElementsOfSecondaryProfiles.Count
            Set oProfile = pElementsOfSecondaryProfiles(i)
            Let sKey = pElementsOfSecondaryProfiles.GetKey(oProfile)

            Set oLineAxisPortExtractor = pGCMemberFactory.LineAxisPortExtractor.Create(Nothing, oProfile)
            oLineAxisPortExtractor.Evaluate
            
            pGeometricConstruction.ControlledInputs("Port2").Add oLineAxisPortExtractor.ControlledInputs("Port")(1), sKey
        Next
    End If
    
    ' retrieve a collection of the axes of the primary members and the computed common node and check if all Members go through the node
    Call DebugValue("ElementsOfPrimaryProfiles.Count", pElementsOfPrimaryProfiles.Count)
    Call DebugValue("ElementsOfSecondaryProfiles.Count", pElementsOfSecondaryProfiles.Count)
    If pElementsOfPrimaryProfiles.Count > 1 Then
        Dim pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode As IJElements
        Set pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode = MemberParts_GetElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode(pElementsOfPrimaryProfiles, Nothing)
    ElseIf pElementsOfSecondaryProfiles.Count > 0 Then
        Call DebugMsg("create a mixed collection with 1 primary and 1 secondary Members")
        Dim pElementsOfMixedProfiles As IJElements: Set pElementsOfMixedProfiles = New JObjectCollection
        Call pElementsOfMixedProfiles.Add(pElementsOfPrimaryProfiles(1), "1")
        Call pElementsOfMixedProfiles.Add(pElementsOfSecondaryProfiles(1), "2")
        
        ' use the 2 Members to detect a node
        Set pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode = MemberParts_GetElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode(pElementsOfMixedProfiles, Nothing)
        If pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode.Count = 0 Then Err.Raise COMMON_NODE_NOT_FOUND
        
        ' purge the list by removing the entries added for the extra secondary Member
        For i = pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode.Count To 1 Step -1
            Let sKey = GetKey(pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode, i)
            If Mid(sKey, 1, 1) = "2" Then
                Call pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode.Remove(pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode.GetKey(pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode(2)))
            End If
        Next
        Call DebugValue("Purged list of ElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode", pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode)
    Else
        Err.Raise MISSING_MANDATORY_INPUT
    End If
    
    ' retrieve node
    Dim pPositionOfNode As IJDPosition
    If True Then
        If pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode.Count = 0 Then Err.Raise COMMON_NODE_NOT_FOUND
        Set pPositionOfNode = pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode("0")
        pGeometricConstructionMacro.Output("Node", 1) = Point_FromPosition(pPositionOfNode)
        Call DebugValue("Adding 'Node'output", pPositionOfNode)
    End If
    
    ' retrieve a collection of the axes of secondary members and check if all Members go through the node
    Dim pElementsOfLinesOfAxesOfSecondaryMembersAndPositionOfCommonNode As IJElements
    If pElementsOfSecondaryProfiles.Count > 0 Then
        Set pElementsOfLinesOfAxesOfSecondaryMembersAndPositionOfCommonNode = MemberParts_GetElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode(pElementsOfSecondaryProfiles, pPositionOfNode)
        If pElementsOfLinesOfAxesOfSecondaryMembersAndPositionOfCommonNode.Count = 0 Then Err.Raise COMMON_NODE_NOT_SHARED
    End If
    
    ' fill an array with the lines of the axes of the primary Members
    Dim pLinesOfAxesOfPrimaryMembers() As IJLine
    If True Then
        ' skip the node at the end of the collection
        ReDim pLinesOfAxesOfPrimaryMembers(1 To pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode.Count - 1)
        For i = 1 To pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode.Count - 1
            Set pLinesOfAxesOfPrimaryMembers(i) = pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode(i)
        Next
    End If

    ' determine a line normal to the APS
    Dim pLineOfNormalToAPS As IJLine: Set pLineOfNormalToAPS = Nothing
    If True Then
        ' try to get it as the normal to the axes of the primary Members
        If True Then
            Set pLineOfNormalToAPS = GetLineAtNormalToLines(GetGCGeomFactory(), GetGCGeomFactory2(), pPOM, pLinesOfAxesOfPrimaryMembers)
            If Not pLineOfNormalToAPS Is Nothing Then
                Call DebugValue("LineOfNormalToAPS (rtrieved as normal to lines)", pLineOfNormalToAPS)
            Else
                If pElementsOfPrimaryProfiles.Count > 2 Then Err.Raise MEMBER_AXES_NOT_COPLANAR
            End If
        
        End If
        
        ' try to use the first secondary Member to define a z direction
        If pLineOfNormalToAPS Is Nothing And pElementsOfSecondaryProfiles.Count > 0 Then
            ' build a CS with the first primary Member and the first column Member and use the axis of the CS
            Dim oCSByLines As Object
            Set oCSByLines = GetGCGeomFactory2().CSByLines(Nothing, _
                                                           pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode(1), _
                                                           pElementsOfLinesOfAxesOfSecondaryMembersAndPositionOfCommonNode(1), _
                                                           Nothing, _
                                                           GCYZ, GCDirect)
            Set pLineOfNormalToAPS = GetGCGeomFactory().LineFromCS.PlaceGeometry(Nothing, oCSByLines, GCZAxis, 1)
            Call DebugValue("LineOfNormalToAPS (retrieved from first primary Memberand first secondary Member)", pLineOfNormalToAPS)
        End If
        
        ' retrieve the line as the z axis of the CS from the first primary Member
        If pLineOfNormalToAPS Is Nothing Then
            Dim oCSFromMember As Object
            Set oCSFromMember = pGCMemberFactory.CSFromMember.PlaceGeometry(Nothing, pGeometricConstruction.Inputs(sPRIMARY_PROFILES)(1), Point_FromPosition(pPositionOfNode))
            Set pLineOfNormalToAPS = GetGCGeomFactory().LineFromCS.PlaceGeometry(Nothing, oCSFromMember, GCZAxis, 1)
            Call DebugValue("LineOfNormalToAPS (retrieved from the CS from the first primary Member)", pLineOfNormalToAPS)
        End If

        ' verify that all the primary Members are orthogonal to the line normal to the APS, if deducted from a secondary Member
        If pElementsOfPrimaryProfiles.Count > 0 Then
            For i = 1 To pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode.Count - 1
                If Not AreVectorsPerpendicular(Vector_FromLine(pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode(i)), _
                                               Vector_FromLine(pLineOfNormalToAPS)) Then Err.Raise MEMBER_AXES_NOT_ORTHOGONAL
            Next
        End If
    End If
             
    ' order lines
    Dim lOrderedIndexes() As Long
    Dim dOrderedAngles() As Double
    Call Lines_OrderAroundPointAndNormal(pLinesOfAxesOfPrimaryMembers, pPositionOfNode, Vector_FromLine(pLineOfNormalToAPS), lOrderedIndexes, dOrderedAngles)
    
    ' create outputs
    Dim pLineOfPreviousLine As IJLine
    For i = 1 To UBound(pLinesOfAxesOfPrimaryMembers, 1)
        Dim j As Integer: Let j = lOrderedIndexes(i)
        Dim sExtendedKey As String: Let sExtendedKey = GetKey(pElementsOfLinesOfAxesOfPrimaryMembersAndPositionOfCommonNode, j)
        ' reorient line
        Dim pPositionOfStartPoint As IJDPosition: Set pPositionOfStartPoint = Position_FromLine(pLinesOfAxesOfPrimaryMembers(j), 0)
        Dim pPositionOfEndPoint As IJDPosition: Set pPositionOfEndPoint = Position_FromLine(pLinesOfAxesOfPrimaryMembers(j), 1)
        Dim pLineOfCurrentLine As IJLine
        If pPositionOfNode.DistPt(pPositionOfStartPoint) < TOLERANCE Then
            Set pLineOfCurrentLine = pLinesOfAxesOfPrimaryMembers(j)
        Else
            Set pLineOfCurrentLine = Line_FromPositions(Nothing, pPositionOfEndPoint, pPositionOfStartPoint)
        End If
        
        ' in the case of 2 stiffeners (or edge-reinforcements) sharing the same root landing curve, but bounded by the same boundary
        Dim bIsLineDuplicate As Boolean: Let bIsLineDuplicate = False
        If i > 1 Then
            If Position_FromLine(pLineOfCurrentLine, 0).DistPt(Position_FromLine(pLineOfPreviousLine, 0)) < EPSILON _
            And Position_FromLine(pLineOfCurrentLine, 1).DistPt(Position_FromLine(pLineOfPreviousLine, 1)) < EPSILON Then
                Let bIsLineDuplicate = True
            End If
        End If
        
        If Not bIsLineDuplicate Then
            pGeometricConstructionMacro.Output("Axis", sExtendedKey) = pLineOfCurrentLine
            Set pLineOfPreviousLine = pLineOfCurrentLine
            Call DebugValue("Adding 'Axis." + sExtendedKey + "' output", pLinesOfAxesOfPrimaryMembers(j))
        End If
    Next
    Set oCSByLines = GetGCGeomFactory2().CSByLines(Nothing, pLineOfNormalToAPS, pLinesOfAxesOfPrimaryMembers(lOrderedIndexes(1)), Nothing, GCZX, GCDirect)
    
    Dim iDirection As Integer: Let iDirection = GeometricConstruction_GetDirection(pGeometricConstruction, oCSByLines)
    If pGeometricConstruction.Parameter("Support") = 1 Then Let iDirection = -iDirection
    If iDirection < 0 Then
        Set oCSByLines = GetGCGeomFactory2().CSByCS(Nothing, oCSByLines, Nothing, Nothing, GCXY, GCIndirect, GCNear)
    End If
    
    pGeometricConstructionMacro.Output("CoordinateSystem", 1) = oCSByLines
    Call DebugValue("Adding coordinate system as output", oCSByLines)
    
    ' extract z axis from CS
    Dim oZAxis As Object
    Set oZAxis = GetGCGeomFactory().LineFromCS.PlaceGeometry(Nothing, oCSByLines, GCZAxis, 1)
    
    pGeometricConstructionMacro.Output("Normal", 1) = oZAxis
    Call DebugValue("Adding 'Normal' output", pLineOfNormalToAPS)

   
    Call DebugOut
    Exit Sub
ErrorHandler:
    Call GCProcessError(pGeometricConstruction, , Err.Number)
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
        'If Mid(sName, 1, Len("MemberPart")) = "MemberPart" Then Call GeometricConstruction_MigrateInputs(MyGC, pMigrateHelper, sName, pPositionOfNode)
        If InStr(1, sName, "MemberPart", vbTextCompare) > 0 Then Call GeometricConstruction_MigrateInputs(MyGC, pMigrateHelper, sName, pPositionOfNode)
    Next

    Call DebugOut
End Sub
Function GeometricConstructionMacro_GetIndexFromName(pGeometricConstructionDefinition As IJGeometricConstructionDefinition, sNameOfInput As String) As Long
    Call DebugIn(Method("GeometricConstructionMacro_GetIndexFromName"))
    Call DebugInput("Name", sNameOfInput)
    
    ' prepare the result
    Dim lIndex As Long: Let lIndex = -1
    
    Dim i As Integer
    For i = 1 To pGeometricConstructionDefinition.InputCount
        Dim sName As String, sPrompt As String, sFilter As String, lNinConnected As Long, lMaxConnected As Long, sComputeIIDs As String
        Call pGeometricConstructionDefinition.GetInputInfoByIndex(i, sName, sPrompt, sFilter, lNinConnected, lMaxConnected, sComputeIIDs)
        If sName = sNameOfInput Then
            Let lIndex = i
            Exit For
        End If
    Next
    
    ' return result
    Let GeometricConstructionMacro_GetIndexFromName = lIndex
    Call DebugOutput("Index", GeometricConstructionMacro_GetIndexFromName)
    Call DebugOut
End Function
Function GetKey(pElements As IJElements, iIndex As Integer) As String
    Dim sKey As String: Let sKey = pElements.GetKey(pElements(iIndex))
'    If sKey = "" Then sKey = CStr(iIndex)
    Let GetKey = sKey
End Function
Private Function GeometricConstruction_GetDirection(pGeometricConstruction As IJGeometricConstruction, oCS As Object) As Integer
     Call DebugIn(Method("GeometricConstruction_GetDirection"))
   ' prepare result
    Dim iDirection As Integer: Let iDirection = 1
    
    Dim pLocalCoordinateSystem As IJLocalCoordinateSystem: Set pLocalCoordinateSystem = oCS
    Dim pVectorOfZAxis As IJDVector: Set pVectorOfZAxis = pLocalCoordinateSystem.ZAxis
    Call DebugValue("VectorOfZAxis", pVectorOfZAxis)
    Dim pVectorOfXAxisOfReference As IJDVector
    Dim pVectorOfYAxisOfReference As IJDVector
    Dim pVectorOfZAxisOfReference As IJDVector
    If pGeometricConstruction.Inputs(sCOORDINATE_SYSTEM).Count > 0 Then
        Dim pLocalCoordinateSystemOfReference As IJLocalCoordinateSystem: Set pLocalCoordinateSystemOfReference = pGeometricConstruction.Input(sCOORDINATE_SYSTEM)
        Set pVectorOfXAxisOfReference = pLocalCoordinateSystemOfReference.XAxis
        Set pVectorOfYAxisOfReference = pLocalCoordinateSystemOfReference.YAxis
        Set pVectorOfZAxisOfReference = pLocalCoordinateSystemOfReference.ZAxis
    Else
        Set pVectorOfXAxisOfReference = New DVector: Call pVectorOfXAxisOfReference.Set(1, 0, 0)
        Set pVectorOfYAxisOfReference = New DVector: Call pVectorOfYAxisOfReference.Set(0, 1, 0)
        Set pVectorOfZAxisOfReference = New DVector: Call pVectorOfZAxisOfReference.Set(0, 0, 1)
    End If
    Call DebugValue("VectorOfXAxisOfReference", pVectorOfXAxisOfReference)
    Call DebugValue("VectorOfYAxisOfReference", pVectorOfYAxisOfReference)
    Call DebugValue("VectorOfZAxisOfReference", pVectorOfZAxisOfReference)
    
    Dim dDotWithXAxisOfReference As Double: Let dDotWithXAxisOfReference = pVectorOfZAxis.Dot(pVectorOfXAxisOfReference)
    Dim dDotWithYAxisOfReference As Double: Let dDotWithYAxisOfReference = pVectorOfZAxis.Dot(pVectorOfYAxisOfReference)
    Dim dDotWithZAxisOfReference As Double: Let dDotWithZAxisOfReference = pVectorOfZAxis.Dot(pVectorOfZAxisOfReference)
    Call DebugValue("DotWithXAxisOfReference", dDotWithXAxisOfReference)
    Call DebugValue("DotWithYAxisOfReference", dDotWithYAxisOfReference)
    Call DebugValue("DotWithZAxisOfReference", dDotWithZAxisOfReference)
    
    If Abs(dDotWithXAxisOfReference) > 0.707 Then iDirection = Sgn(dDotWithXAxisOfReference)
    If Abs(dDotWithYAxisOfReference) > 0.707 Then iDirection = Sgn(dDotWithYAxisOfReference)
    If Abs(dDotWithZAxisOfReference) > 0.707 Then iDirection = Sgn(dDotWithZAxisOfReference)

    ' return result
    Call DebugValue("Direction", iDirection)
    Let GeometricConstruction_GetDirection = iDirection
    
    Call DebugOut
End Function

