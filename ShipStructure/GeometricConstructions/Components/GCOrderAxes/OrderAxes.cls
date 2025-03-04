VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderAxes"
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
End Enum

Const m_sModule = "GCOrderAxes.OrderAxes"
Const sORTHOGONAL_MEMBER_PART = "OrthogonalMemberPart"
Const sSUPPORTING_MEMBER_PART = "SupportingMemberPart"

Implements IJGeometricConstructionDefinitionService
Private Property Get Source() As String
    Let Source = "GCOrderAxes.OrderAxes"
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
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Initialize: OrderAxes"))
    
    ' located inputs
    Call pGeometricConstructionDefinition.AddInput("MemberParts", "Select n coplanar Members in consecutive counterclockwise or clockwise order", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 1, 8, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput(sORTHOGONAL_MEMBER_PART, "Select an orthogonal Member", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput(sSUPPORTING_MEMBER_PART, "Select a supporting Member", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
        
    ' controlled inputs
    Call pGeometricConstructionDefinition.AddControlledInput("Port")
    
    ' code listed parameters
    Call pGeometricConstructionDefinition.AddParameter("Support", "Side", GCCodeList, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Back", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Front", 2)
    
    ' error codes
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MANDATORY_INPUT, "MissingMandatoryInput", "At least 2 members are required")
    Call pGeometricConstructionDefinition.AddErrorValue(COMMON_NODE_NOT_FOUND, "CommonNodeNotFound", "Common node not found")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXIS_TOO_SMALL, "MemberAxisTooSmall", "At least 1 member axis is too small")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXES_COLINEAR, "MemberAxesColinear", "Member axes colinear")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXES_NOT_COPLANAR, "MemberAxesNotCoplanar", "Member axes not coplanar")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_FACES_NOT_PARALLEL, "MemberFacesNotParallel", "Member faces not parallel")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AXES_NOT_ORTHOGONAL, "MemberAxesNotOrthogonal", "Member axes not perpendicular to the orthogonal member")
    
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
    
    ' use GC as a GCmacro
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGeometricConstruction
    
    ' retrieve info about the incoming members
    Dim pElementsOfPrimaryMemberParts As IJElements: Set pElementsOfPrimaryMemberParts = pGeometricConstruction.Inputs("MemberParts")
    Dim iCountOfPrimaryMemberParts As Integer: Let iCountOfPrimaryMemberParts = pElementsOfPrimaryMemberParts.Count
    Dim sKeysOfPrimaryMemberParts() As String: If iCountOfPrimaryMemberParts > 0 Then Let sKeysOfPrimaryMemberParts = Elements_GetKeys(pElementsOfPrimaryMemberParts)
    If iCountOfPrimaryMemberParts < 1 Then Err.Raise MISSING_MANDATORY_INPUT
    
    ' extract LineAxisPort of each incoming Member
    Dim pGCMemberFactory As IJGCMemberFactory: Set pGCMemberFactory = CreateObject("GCSPSDefinitions.GCMemberFactory")
    pGeometricConstruction.ControlledInputs("Port").Clear
    Dim i As Integer
    For i = 1 To iCountOfPrimaryMemberParts
        Dim pGeometricConstructionEntitiesFactory As IJGeometricConstructionEntitiesFactory: Set pGeometricConstructionEntitiesFactory = New GeometricConstructionEntitiesFactory
        Dim oLineAxisPortExtractor As IJGeometricConstruction
        Set oLineAxisPortExtractor = pGCMemberFactory.LineAxisPortExtractor.Create(Nothing, pElementsOfPrimaryMemberParts(i))
        oLineAxisPortExtractor.Evaluate
        pGeometricConstruction.ControlledInputs("Port").Add oLineAxisPortExtractor.ControlledInputs("Port")(1)
    Next
            
    ' retrieve the optional orthogonal member
    Dim oOrthogonalMemberPart As Object:
    If GeometricConstructionMacro_GetIndexFromName(pGeometricConstruction.definition, sORTHOGONAL_MEMBER_PART) > 0 Then
        If pGeometricConstruction.Inputs(sORTHOGONAL_MEMBER_PART).Count = 1 Then
            Set oOrthogonalMemberPart = pGeometricConstruction.Input(sORTHOGONAL_MEMBER_PART)
        End If
    End If
    
    ' retrieve the optional supporting member
    Dim oSupportingMemberPart As Object:
    If GeometricConstructionMacro_GetIndexFromName(pGeometricConstruction.definition, sSUPPORTING_MEMBER_PART) > 0 Then
        If pGeometricConstruction.Inputs(sSUPPORTING_MEMBER_PART).Count = 1 Then
            Set oSupportingMemberPart = pGeometricConstruction.Input(sSUPPORTING_MEMBER_PART)
        End If
   End If
    
    ' compute common node
    Dim pPositionOfNode As IJDPosition
    If True Then
        If Not oOrthogonalMemberPart Is Nothing Or Not oSupportingMemberPart Is Nothing Then
            Dim oOrthogonalOrSupportingMember As Object
            If Not oOrthogonalMemberPart Is Nothing Then Set oOrthogonalOrSupportingMember = oOrthogonalMemberPart
            If Not oSupportingMemberPart Is Nothing Then Set oOrthogonalOrSupportingMember = oSupportingMemberPart
            
            ' compute node as point on the orthogonal/supporting member
            Set pPositionOfNode = MemberParts_GetPositionAlong(oOrthogonalOrSupportingMember, pElementsOfPrimaryMemberParts(1))
            
            ' verify that the primary MemberParts produce the same node
            If iCountOfPrimaryMemberParts > 1 Then
                Dim pPositionOfNode1 As IJDPosition: Set pPositionOfNode1 = MemberParts_GetPositionOfCommonEnd(pElementsOfPrimaryMemberParts)
                If pPositionOfNode.DistPt(pPositionOfNode1) > EPSILON Then Err.Raise COMMON_NODE_NOT_FOUND
            End If
        Else
            Set pPositionOfNode = MemberParts_GetPositionOfCommonEnd(pElementsOfPrimaryMemberParts)
        End If
    End If
    If pPositionOfNode Is Nothing Then Err.Raise COMMON_NODE_NOT_FOUND
    Call DebugValue("PositionOfNode", pPositionOfNode)
    
    ' retrieve lines of member axes
     Dim pLinesOfMemberAxes() As IJLine: ReDim pLinesOfMemberAxes(1 To iCountOfPrimaryMemberParts)
     Dim iCountOfLinesOfMemberAxes As Integer: Let iCountOfLinesOfMemberAxes = 0
     Dim pLineOfOrthogonalMemberAxis As IJLine
     If True Then
        If Not oSupportingMemberPart Is Nothing Then
            ' add the supporting MemberPart as first Member
            ReDim pLinesOfMemberAxes(1 To iCountOfPrimaryMemberParts + 2)
            Dim pPositionsAtEndsAndJointsOfMember() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember = MemberPart_GetPositionsAtJointsAndEnds(oSupportingMemberPart)

            Set pLinesOfMemberAxes(1) = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(0), pPositionOfNode)
            If pLinesOfMemberAxes(1).Length < EPSILON Then Err.Raise MEMBER_AXIS_TOO_SMALL
            
            Set pLinesOfMemberAxes(2) = Line_FromPositions(Nothing, pPositionOfNode, pPositionsAtEndsAndJointsOfMember(1))
            If pLinesOfMemberAxes(2).Length < EPSILON Then Err.Raise MEMBER_AXIS_TOO_SMALL
            iCountOfLinesOfMemberAxes = 2
        End If
        
        For i = 1 To iCountOfPrimaryMemberParts
            iCountOfLinesOfMemberAxes = iCountOfLinesOfMemberAxes + 1
            Set pLinesOfMemberAxes(iCountOfLinesOfMemberAxes) = MemberPart_GetLineOfMemberAxisAtNode(pElementsOfPrimaryMemberParts(i), pPositionOfNode)
            If pLinesOfMemberAxes(iCountOfLinesOfMemberAxes).Length < EPSILON Then Err.Raise MEMBER_AXIS_TOO_SMALL
        Next
        
        If Not oOrthogonalMemberPart Is Nothing Then
            Set pLineOfOrthogonalMemberAxis = MemberPart_GetLineOfMemberAxisThroughNode(oOrthogonalMemberPart, pPositionOfNode)
            
            ' check if all primary MemberParts are perpendicular to the orthogonal MemberPart
            For i = 1 To iCountOfPrimaryMemberParts
                If Not AreVectorsPerpendicular(Vector_FromLine(pLinesOfMemberAxes(i)), Vector_FromLine(pLineOfOrthogonalMemberAxis)) Then Err.Raise MEMBER_AXES_NOT_ORTHOGONAL
            Next
        End If
    End If
    
    Dim pLineOfNormal As IJLine
    If True Then
        ' compute line normal to lines
        Set pLineOfNormal = GetLineAtNormalToLines(GetGCGeomFactory(), GetGCGeomFactory2(), pPOM, pLinesOfMemberAxes)
        
        If pLineOfNormal Is Nothing Then
            If iCountOfPrimaryMemberParts = 1 And Not pLineOfOrthogonalMemberAxis Is Nothing Then
                Set pLineOfNormal = pLineOfOrthogonalMemberAxis
            ElseIf iCountOfPrimaryMemberParts = 2 Then
                'Err.Raise MEMBER_AXES_COLINEAR
                Dim pCSFromMember As IJDCoordinateSystem
                Set pCSFromMember = pGCMemberFactory.CSFromMember.PlaceGeometry(Nothing, pGeometricConstruction.Inputs("MemberParts")(1), Point_FromPosition(pPositionOfNode))
                
                Set pLineOfNormal = GetGCGeomFactory().LineFromCS.PlaceGeometry(Nothing, pCSFromMember, GCZAxis, 1)
            Else
                Err.Raise MEMBER_AXES_NOT_COPLANAR
            End If
        End If
    End If
        
    ' order lines
    Dim lOrderedIndexes() As Long
    Dim dOrderedAngles() As Double
    Call Lines_OrderAroundPointAndNormal(pLinesOfMemberAxes, pPositionOfNode, Vector_FromLine(pLineOfNormal), lOrderedIndexes, dOrderedAngles)
    
    ' create outputs
    For i = 1 To iCountOfLinesOfMemberAxes
        Dim j As Integer: Let j = lOrderedIndexes(i)
        Dim sKey As String
        If Not oSupportingMemberPart Is Nothing Then
            If j = 1 Then
                Let sKey = "0"
            ElseIf j = 2 Then
                Let sKey = CStr(iCountOfPrimaryMemberParts + 1)
            Else
                sKey = sKeysOfPrimaryMemberParts(j - 2)
            End If
        Else
            sKey = sKeysOfPrimaryMemberParts(j)
        End If
        pGeometricConstructionMacro.Output("Axis", sKey) = pLinesOfMemberAxes(j)
    Next
    pGeometricConstructionMacro.Output("Normal", 1) = pLineOfNormal
    pGeometricConstructionMacro.Output("Node", 1) = Point_FromPosition(pPositionOfNode)
    
    Dim oCSByLines As Object: Set oCSByLines = GetGCGeomFactory2().CSByLines(Nothing, pLineOfNormal, pLinesOfMemberAxes(lOrderedIndexes(1)), Nothing, GCZX, GCDirect)
    pGeometricConstructionMacro.Output("CoordinateSystem", 1) = oCSByLines
   
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

