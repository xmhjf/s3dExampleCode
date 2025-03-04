Attribute VB_Name = "Common"
Public Const EPSILON As Double = 0.00001
Public Const TOLERANCE As Double = 0.001
Public Const PI = 3.14159265358979
Public Const sCOORDINATE_SYSTEM = "CoordinateSystem"
Public m_pGCGeomFactory As IJGCGeomFactory
Public m_pGCGeomFactory2 As IJGCGeomFactory2

Public Function GetGCGeomFactory() As IJGCGeomFactory
    If m_pGCGeomFactory Is Nothing Then
        Set m_pGCGeomFactory = CreateObject("GCCMNSTRDefinitions.GCGeomFactory")
    End If
    Set GetGCGeomFactory = m_pGCGeomFactory
End Function
Public Function GetGCGeomFactory2() As IJGCGeomFactory2
    If m_pGCGeomFactory2 Is Nothing Then
        Set m_pGCGeomFactory2 = CreateObject("GCCMNSTRDefinitions.GCGeomFactory")
    End If
    Set GetGCGeomFactory2 = m_pGCGeomFactory2
End Function
Public Function Position_AtCurvesIntersection(pCurve1 As IJCurve, pCurve2 As IJCurve) As IJDPosition
    ' initialize result
    Set Position_AtCurvesIntersection = Nothing
    
    ' compute all possible positions of intersection points
    Dim pElementsOfPositions As IJElements
    If True Then
        ' compute positions as an array of doubles
        Dim dDoubles() As Double
        Dim lNumIntersections As Long: Let lNumIntersections = 0
        If True Then
            Dim lNumOverlaps As Long
            On Error Resume Next
            Call pCurve1.Intersect(pCurve2, lNumIntersections, dDoubles, lNumOverlaps, ISECT_UNKNOWN)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
        If lNumIntersections = 0 Then Exit Function
        
        ' transform an array of doubles into a collection of positions
        Set pElementsOfPositions = Positions_FromDoubles(dDoubles, lNumIntersections)
        
        If pElementsOfPositions Is Nothing Then Exit Function
        If pElementsOfPositions.Count = 0 Then Exit Function
    End If
        
    ' return result
    Set Position_AtCurvesIntersection = pElementsOfPositions.Item(1)
End Function

Public Function Positions_FromDoubles(dDoubles() As Double, lCount As Long) As IJElements
    ' initialize result
    Set Positions_FromDoubles = Nothing
    
    ' create collection of positions
    Dim pElementsOfPositions As IJElements: Set pElementsOfPositions = New JObjectCollection
    
    ' fill the collection of positions
    Dim i As Integer
    For i = 1 To lCount
        Dim pPosition As IJDPosition: Set pPosition = New DPosition
        Dim j As Integer: Let j = 3 * (i - 1)
        Call pPosition.Set(dDoubles(j), dDoubles(j + 1), dDoubles(j + 2))
        Call pElementsOfPositions.Add(pPosition)
    Next
    
    ' return result
    Set Positions_FromDoubles = pElementsOfPositions
End Function

Public Function Point_FromPosition(pPosition As IJDPosition) As IJPoint
    ' create new position
    Dim pPoint As IJPoint
    Set pPoint = New Point3d
    
    ' retrieve coordinates
    Dim x As Double, y As Double, z As Double
    Call pPosition.Get(x, y, z)
    
    ' set coordinates
    Call pPoint.SetPoint(x, y, z)
    
    ' return result
    Set Point_FromPosition = pPoint
End Function

Public Function Line_FromPositions(pPOM As IJDPOM, pPositionOfStartPoint As IJDPosition, pPositionOfEndPoint As IJDPosition) As IJLine
    ' create new line
    Set GetGeometryFactory = New GeometryFactory
    Set Line_FromPositions = GetGeometryFactory.Lines3d.CreateBy2Points(pPOM, _
        pPositionOfStartPoint.x, pPositionOfStartPoint.y, pPositionOfStartPoint.z, _
        pPositionOfEndPoint.x, pPositionOfEndPoint.y, pPositionOfEndPoint.z)
End Function

'Method to Update the oOrderAxes21 with Stiffener Parts Axis

Public Function UpdateOrderAxeswithStiffenerParts(ByRef oOrderAxes21 As IJGeometricConstruction, ByVal pGeometricConstruction As IJGeometricConstruction)
    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oLine() As IJCurve
    Dim iInputCount As Integer
    Dim iDirection As Integer
    
    'Get Inputs Count
    iInputCount = oOrderAxes21.Inputs("PrimaryProfiles").Count
    
    ReDim oLine(1 To iInputCount) As IJCurve
    Dim oSupportingLine As IJLine
    Dim iCount As Integer
    
    Dim oSupportedElements As IJElements
    Set oSupportedElements = New JObjectCollection
    
    'Get all the Supported Stiffeners
    For iCount = 2 To iInputCount
        oSupportedElements.Add pGeometricConstruction.Inputs("PrimaryProfile" & iCount).Item(1), CStr(iCount)
    Next
    
    'Get Direction
    iDirection = pGeometricConstruction.Parameter("Direction")
        
    Dim oExtractPorts1 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtractPorts1 = oGCFactory.CreateEntity("ExtractPorts", Nothing, "0001-ExtractPorts")
    oExtractPorts1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "2"
    oExtractPorts1.Inputs("ZAxis").Add oOrderAxes21.Output("Normal", 1), "1"
    oExtractPorts1.Parameter("Support") = 2
    oExtractPorts1.Parameter("Debug") = 0
    oExtractPorts1.Evaluate
    
    'Get Supporting Stiffener Part Axis
    Set oSupportingLine = oExtractPorts1.Output("MemberAxis", 1)
    
    
    Dim oIntersection As IJDPosition
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    
    Set oStartPos = New DPosition
    Set oEndPos = New DPosition
    
    Dim x As Double, y As Double, z As Double
    Dim ex As Double, ey As Double, ez As Double
    
    'Get Supported Stiffener Part Axis
    For iCount = 1 To oSupportedElements.Count
    
        Dim oExtractPorts2 As SP3DGeometricConstruction.GeometricConstruction
        Set oExtractPorts2 = oGCFactory.CreateEntity("ExtractPorts", Nothing, "0002-ExtractPorts")
        oExtractPorts2.Inputs("MemberPart").Add oSupportedElements.Item(iCount), "2"
        oExtractPorts2.Inputs("ZAxis").Add oOrderAxes21.Output("Normal", 1), "1"
        oExtractPorts2.Parameter("Support") = 2
        oExtractPorts2.Parameter("Debug") = 0
        oExtractPorts2.Evaluate
        
        Dim oCurveLeft As IJCurve, oCurveRight As IJCurve
        Dim oStartPoint As IJPoint, oEndPoint As IJPoint
        Dim dMinDist As Double
        
        Set oCurveLeft = oExtractPorts2.Output("CurveLeft", 1)
        Set oCurveRight = oExtractPorts2.Output("CurveRight", 1)
        
        oCurveLeft.EndPoints x, y, z, ex, ey, ez
        
        oStartPos.Set x, y, z
        oEndPos.Set ex, ey, ez
        
        oCurveRight.EndPoints x, y, z, ex, ey, ez
                        
        Set oStartPoint = New Point3d
        Set oEndPoint = New Point3d
                        
        oStartPoint.SetPoint (x + oStartPos.x) / 2, (y + oStartPos.y) / 2, (z + oStartPos.z) / 2
        oEndPoint.SetPoint (ex + oEndPos.x) / 2, (ey + oEndPos.y) / 2, (ez + oEndPos.z) / 2
        
        Set oLine(iCount) = Line_FromPositions(Nothing, Position_FromPoint(oStartPoint), Position_FromPoint(oEndPoint))
        
        Set oExtractPorts2 = Nothing
    Next
    
    For iCount = 1 To oSupportedElements.Count
        'Get Intersection of Supporting and Supported
        Set oIntersection = Position_AtCurvesIntersection(oSupportingLine, oLine(iCount))
        
        'If not intersecting then check with infinate Line
        If oIntersection Is Nothing Then
            Dim oInfinateLine As IJLine
            Set oInfinateLine = New Line3d
            
            oLine(iCount).EndPoints x, y, z, ex, ey, ez
            
            oInfinateLine.SetStartPoint x, y, z
            oInfinateLine.SetEndPoint ex, ey, ez
            oInfinateLine.Infinite = True
            
            'Get Intersection
            Set oIntersection = Position_AtCurvesIntersection(oSupportingLine, oInfinateLine)
            
            If Not oIntersection Is Nothing Then Exit For
        Else
            Exit For
        End If
    Next
        
    'If no Intersection, then Exit
    If oIntersection Is Nothing Then Exit Function
        
    'Update the Lines from Intersection Position to other End
    For iCount = 1 To oSupportedElements.Count
        
        Set oStartPos = New DPosition
        Set oEndPos = New DPosition

        oLine(iCount).EndPoints x, y, z, ex, ey, ez
        
        oStartPos.Set x, y, z
        oEndPos.Set ex, ey, ez
        
        Set oLine(iCount) = Nothing
        
        'Create line from Intersection Position to Far End of Stiffener part Axis
        If Abs(oIntersection.DistPt(oStartPos)) > Abs(oIntersection.DistPt(oEndPos)) Then
            Set oLine(iCount) = Line_FromPositions(Nothing, oIntersection, oStartPos)
        Else
            Set oLine(iCount) = Line_FromPositions(Nothing, oIntersection, oEndPos)
        End If
    Next
        
    Set oStartPos = New DPosition
    Set oEndPos = New DPosition
    
    oSupportingLine.GetStartPoint x, y, z
    oStartPos.Set x, y, z
    oSupportingLine.GetEndPoint ex, ey, ez
    oEndPos.Set ex, ey, ez
           
    Dim pLinesOfAxesOfProfileParts() As IJLine
    Dim pElementsOfLinesOfAxes As IJElements
    Set pElementsOfLinesOfAxes = New JObjectCollection
    
    'Get lines from Intersection Position to Start/Ends of Supporting ProfilePart Axis
    pElementsOfLinesOfAxes.Add Line_FromPositions(Nothing, oIntersection, oStartPos)
    pElementsOfLinesOfAxes.Add Line_FromPositions(Nothing, oIntersection, oEndPos)
    
    For iCount = 1 To oSupportedElements.Count
        
        oLine(iCount).EndPoints x, y, z, ex, ey, ez
        oStartPos.Set x, y, z
        oEndPos.Set ex, ey, ez
        
        If oStartPos.DistPt(oIntersection) <= EPSILON Or oEndPos.DistPt(oIntersection) <= EPSILON Then
            pElementsOfLinesOfAxes.Add oLine(iCount)
        ElseIf oLine(iCount).IsPointOn(oIntersection.x, oIntersection.y, oIntersection.z) Then
            pElementsOfLinesOfAxes.Add Line_FromPositions(Nothing, oIntersection, oStartPos)
            pElementsOfLinesOfAxes.Add Line_FromPositions(Nothing, oIntersection, oEndPos)
        End If
    Next

    ReDim pLinesOfAxesOfProfileParts(1 To pElementsOfLinesOfAxes.Count)
    
    For iCount = 1 To pElementsOfLinesOfAxes.Count
        Set pLinesOfAxesOfProfileParts(iCount) = pElementsOfLinesOfAxes.Item(iCount)
    Next

    ' determine a line normal to the APS
    Dim pLineOfNormalToAPS As IJLine: Set pLineOfNormalToAPS = Nothing
    ' try to get it as the normal to the axes of the primary Members
    Set pLineOfNormalToAPS = GetLineAtNormalToLines(GetGCGeomFactory(), GetGCGeomFactory2(), Nothing, pLinesOfAxesOfProfileParts)

    ' order lines
    Dim lOrderedIndexes() As Long
    Dim dOrderedAngles() As Double
    Call Lines_OrderAroundPointAndNormal(pLinesOfAxesOfProfileParts, oIntersection, Vector_FromLine(pLineOfNormalToAPS), lOrderedIndexes, dOrderedAngles)
    
    ' create outputs
    Dim pLineOfPreviousLine As IJLine
    Dim i As Integer
    For i = 1 To UBound(pLinesOfAxesOfProfileParts)
        Dim j As Integer: Let j = lOrderedIndexes(i)
        ' reorient line
        Dim pPositionOfStartPoint As IJDPosition: Set pPositionOfStartPoint = Position_FromLine(pLinesOfAxesOfProfileParts(j), 0)
        Dim pPositionOfEndPoint As IJDPosition: Set pPositionOfEndPoint = Position_FromLine(pLinesOfAxesOfProfileParts(j), 1)
        Dim pLineOfCurrentLine As IJLine
        If oIntersection.DistPt(pPositionOfStartPoint) < TOLERANCE Then
            Set pLineOfCurrentLine = pLinesOfAxesOfProfileParts(j)
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
            oOrderAxes21.Output("Axis", i) = pLineOfCurrentLine
            Set pLineOfPreviousLine = pLineOfCurrentLine
        End If
    Next
    
    Dim oCSByLines As Object
    Set oCSByLines = GetGCGeomFactory2().CSByLines(Nothing, pLineOfNormalToAPS, pLinesOfAxesOfProfileParts(1), Nothing, GCZX, GCDirect)
        
    If iDirection > 1 Then
        Set oCSByLines = GetGCGeomFactory2().CSByCS(Nothing, oCSByLines, Nothing, Nothing, GCXY, GCIndirect, GCNear)
    End If
    
    oOrderAxes21.Output("CoordinateSystem", 1) = oCSByLines
    
    ' extract z axis from CS
    Dim oZAxis As Object
    Set oZAxis = GetGCGeomFactory().LineFromCS.PlaceGeometry(Nothing, oCSByLines, GCZAxis, 1)
    
    oOrderAxes21.Output("Normal", 1) = oZAxis
    oOrderAxes21.Output("Node", 1) = Point_FromPosition(oIntersection)

End Function

Public Function GetLineAtNormalToLines(pGCGeomFactory As IJGCGeomFactory, pGCGeomFactory2 As IJGCGeomFactory2, pPOM As IJDPOM, pLines() As IJLine) As IJLine
    ' prepare result
    Dim pLineOfNormal As IJLine
    
    ' retrieve size of array
    Dim iLines As Integer: Let iLines = UBound(pLines, 1)
    
    ' verify that the lines are not colinear, when only 2 lines
    If iLines > 2 Or Not AreVectorsColinear(Vector_FromLine(pLines(1)), Vector_FromLine(pLines(2))) Then
        ' deduct a CS from all the lines
        Dim pCoordinateSystem As IJDCoordinateSystem: Set pCoordinateSystem = GetCSByLines(pGCGeomFactory, pGCGeomFactory2, pPOM, pLines)
        If Not pCoordinateSystem Is Nothing Then
            ' compute normal to the plane
            Set pLineOfNormal = pGCGeomFactory.LineFromCS.PlaceGeometry(pPOM, pCoordinateSystem, GCZAxis, 1)
        End If
    End If
    
    ' return result
    Set GetLineAtNormalToLines = pLineOfNormal
End Function

Public Function AreVectorsColinear(pVector1 As DVector, pVector2 As DVector) As Boolean
    
    If Abs(Abs(pVector1.Dot(pVector2)) - pVector1.Length() * pVector2.Length()) < EPSILON Then
        AreVectorsColinear = True
    Else
        AreVectorsColinear = False
    End If
End Function

Public Function Vector_FromLine(pLine As IJLine) As IJDVector
    ' get vector
    Dim u As Double, v As Double, w As Double
    Call pLine.GetDirection(u, v, w)
    
    ' create result
    Dim pVector As New DVector
    Call pVector.Set(u, v, w)
    
    ' return result
    Set Vector_FromLine = pVector
End Function
Public Function GetCSByLines(pGCGeomFactory As IJGCGeomFactory, pGCGeomFactory2 As IJGCGeomFactory2, pPOM As IJDPOM, pLines() As IJLine) As IJDCoordinateSystem
    ' prepare result
    Dim pCoordinateSystem As IJDCoordinateSystem
    
    ' compute CS from the 2 first non colinear member axes
    On Error Resume Next
    Set pCoordinateSystem = pGCGeomFactory2.CSByLines(pPOM, pLines(1), pLines(2), Nothing, GCXY, GCDirect)
    'If Err.Number <> 0 Then
    If pCoordinateSystem Is Nothing Then
        Set pCoordinateSystem = pGCGeomFactory2.CSByLines(pPOM, pLines(1), pLines(3), Nothing, GCXY, GCDirect)
    End If
    On Error GoTo 0

    ' retrieve size of array
    Dim iSize As Integer: Let iSize = UBound(pLines, 1)
    
    ' verify that all lines are in the same plane
    If iSize > 2 And Not pCoordinateSystem Is Nothing Then
        ' compute vector normal to the plane defined by the 2 lines
        Dim pVectorOfNormal As IJDVector: Set pVectorOfNormal = Vector_FromLine(pGCGeomFactory.LineFromCS.PlaceGeometry(pPOM, pCoordinateSystem, GCZAxis, 1))
        
        Dim i As Integer
        For i = 2 To iSize ' restart to 2 in case the CS has been built with the lines 1 and 3
            If pVectorOfNormal.Dot(Vector_FromLine(pLines(i))) > EPSILON Then
                Set pCoordinateSystem = Nothing
                Exit For
            End If
        Next
    End If
    
    ' return result
    Set GetCSByLines = pCoordinateSystem
End Function
Sub Lines_OrderAroundPointAndNormal(pLines() As IJLine, pPositionOfOrigin As IJDPosition, pVectorOfNormal As IJDVector, lOrderedIndexes() As Long, dOrderedAngles() As Double)

    ' collect extremity points
    Dim pPositionsOfExtremityPoints() As IJDPosition
    Let pPositionsOfExtremityPoints = GetPositionsOfExtremitiesOfLines(pLines, pPositionOfOrigin)

    ' compute angles with respect to the line betwen the origin and the first point
    Dim dAngles() As Double
    Let dAngles = GetAnglesFromCenterExtremitiesAndNormal(pPositionOfOrigin, _
                                                        pPositionsOfExtremityPoints, _
                                                        pVectorOfNormal)
                                                        
    Call Angles_OrderByAngle(dAngles, dOrderedAngles, lOrderedIndexes)
End Sub

Public Sub Angles_OrderByAngle(dAngles() As Double, ByRef dOrderedAngles() As Double, ByRef lOrderedIndexes() As Long)
    ' retrieve size of the array
    Dim iSize As Integer: Let iSize = UBound(dAngles, 1)
    
    ' remember already processed elements
    Dim bIsAlreadyProcessed() As Boolean: ReDim bIsAlreadyProcessed(1 To iSize)
    Dim i As Integer: For i = 1 To iSize: Let bIsAlreadyProcessed(i) = False: Next
    
    ' initialize the outputs
    ReDim dOrderedAngles(1 To iSize)
    ReDim lOrderedIndexes(1 To iSize)
    
    ' store the first element
    dOrderedAngles(1) = 0
    lOrderedIndexes(1) = 1
    
    Let bIsAlreadyProcessed(1) = True
    
    ' fill the ordered collection
    For i = 2 To iSize
        Dim dAngleMini As Double: Let dAngleMini = 360 ' degrees
        Dim j As Integer: Dim jOfSmallerAngle As Integer
        For j = 2 To iSize
            If bIsAlreadyProcessed(j) = False Then
                If dAngles(j) < dAngleMini Then
                    Let dAngleMini = dAngles(j)
                    Let jOfSmallerAngle = j
                End If
            End If
        Next
        
        ' store the ith element
        dOrderedAngles(i) = dAngles(jOfSmallerAngle)
        lOrderedIndexes(i) = jOfSmallerAngle
        
        Let bIsAlreadyProcessed(jOfSmallerAngle) = True
    Next
End Sub

Public Function Position_FromLine(pLine As IJLine, iIndex As Integer) As IJDPosition
    ' create new position
    Dim pPosition As IJDPosition
    Set pPosition = New DPosition
    
    ' retrieve coordinates
    Dim x As Double, y As Double, z As Double
    If iIndex = 0 Then
        Call pLine.GetStartPoint(x, y, z)
    Else
        Call pLine.GetEndPoint(x, y, z)
    End If
    
    ' set coordinates
    Call pPosition.Set(x, y, z)
    
    ' return result
    Set Position_FromLine = pPosition
End Function

'
' private functions
'
Public Function GetPositionsOfExtremitiesOfLines(pLines() As IJLine, pPositionOfNode As IJDPosition) As IJDPosition()
    ' initialize result
    Dim pPositionsOfExtremityPoints() As IJDPosition: ReDim pPositionsOfExtremityPoints(1 To UBound(pLines, 1))

    Dim i As Integer
    For i = 1 To UBound(pLines, 1)
        Dim pPosition0 As IJDPosition: Set pPosition0 = Position_FromLine(pLines(i), 0)
        Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromLine(pLines(i), 1)

        If pPositionOfNode.DistPt(pPosition0) < EPSILON Then
            Set pPositionsOfExtremityPoints(i) = pPosition1
        ElseIf pPositionOfNode.DistPt(pPosition1) < EPSILON Then
            Set pPositionsOfExtremityPoints(i) = pPosition0
        Else '  processinterrupted
            Exit For
        End If
    Next

    ' return result
    Let GetPositionsOfExtremitiesOfLines = pPositionsOfExtremityPoints
End Function
Public Function GetAnglesFromCenterExtremitiesAndNormal(pPositionOfCenter As IJDPosition, _
                                                 pPositionsOfExtremityPoints() As IJDPosition, _
                                                 pVectorOfNormal As IJDVector) As Double()
                                               
    ' size of input array
    Dim iSize As Integer: Let iSize = UBound(pPositionsOfExtremityPoints, 1)
    
    Dim dAngles() As Double: ReDim dAngles(1 To iSize)
    Let dAngles(1) = 0
    Dim i As Integer
    For i = 2 To iSize
        Let dAngles(i) = GetAngleFromCenterStartEndAndNormal( _
                            pPositionOfCenter, _
                            pPositionsOfExtremityPoints(1), _
                            pPositionsOfExtremityPoints(i), _
                            pVectorOfNormal)
    Next
    
    ' return result
    Let GetAnglesFromCenterExtremitiesAndNormal = dAngles
End Function

Public Function GetAngleFromCenterStartEndAndNormal(pPositionOfCenterPoint As IJDPosition, _
                                             pPositionOfStartPoint As IJDPosition, _
                                             pPositionOfEndPoint As IJDPosition, _
                                             pVectorOfNormal As IJDVector) As Double
        ' compute arc
        Dim pArc As IJArc
        Set GetGeometryFactory = New GeometryFactory
        Set pArc = GetGeometryFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                            pPositionOfCenterPoint.x, _
                            pPositionOfCenterPoint.y, _
                            pPositionOfCenterPoint.z, _
                            pVectorOfNormal.x, _
                            pVectorOfNormal.y, _
                            pVectorOfNormal.z, _
                            pPositionOfStartPoint.x, _
                            pPositionOfStartPoint.y, _
                            pPositionOfStartPoint.z, _
                            pPositionOfEndPoint.x, _
                            pPositionOfEndPoint.y, _
                            pPositionOfEndPoint.z)
        ' return result
        Let GetAngleFromCenterStartEndAndNormal = pArc.SweepAngle / PI * 180
        
End Function



Public Function Position_FromPoint(pPoint As IJPoint) As IJDPosition
    ' create new position
    Dim pPosition As IJDPosition
    Set pPosition = New DPosition
    
    ' retrieve coordinates
    Dim x As Double, y As Double, z As Double
    Call pPoint.GetPoint(x, y, z)
    
    ' set coordinates
    Call pPosition.Set(x, y, z)
    
    ' return result
    Set Position_FromPoint = pPosition
End Function
