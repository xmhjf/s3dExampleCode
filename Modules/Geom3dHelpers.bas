Attribute VB_Name = "Geom3dHelpers"
Option Explicit
' depends on Geom3d (IJPoint, IJLine, ...)
' depends on ResPOM (IJDPOM)
' depends on AutoMathHelpers.bas
' depends on ExtendedDebugSupport.bas
Private m_oGeometryFactory As GeometryFactory
Public Function GetGeometryFactory() As GeometryFactory
    If m_oGeometryFactory Is Nothing Then
        Set m_oGeometryFactory = New GeometryFactory
    End If
    Set GetGeometryFactory = m_oGeometryFactory
End Function
Function Value_ToString(dValue As Double, sText As String) As String
    Let Value_ToString = sText + ": d= " + CStr(dValue)
End Function

'
' processing Point
'
Sub Point_Debug(sMessage As String, pPoint As IJPoint)
    Dim dX As Double, dY As Double, dZ As Double
    Call pPoint.GetPoint(dX, dY, dZ)
    Debug.Print (sMessage + ": X= " + CStr(dX) + ", Y= " + CStr(dY) + ". Z= " + CStr(dZ))
End Sub
Function Position_FromPoint(pPoint As IJPoint) As IJDPosition
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
Function Vector_FromPoint(pPoint As IJPoint) As IJDVector
    ' create new vector
    Dim pVectorNew As IJDVector
    Set pVectorNew = New DVector
    
    ' retrieve coordinates
    Dim x As Double, y As Double, z As Double
    Call pPoint.GetPoint(x, y, z)
    
    ' set coordinates
    Call pVectorNew.Set(x, y, z)

    ' return result
    Set Vector_FromPoint = pVectorNew
End Function
Function Point_FromPosition(pPosition As IJDPosition) As IJPoint
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
Sub Point_UpdateFromPosition(pPoint As IJPoint, pPosition As IJDPosition)
    Call pPoint.SetPoint(pPosition.x, _
                         pPosition.y, _
                         pPosition.z)
End Sub
Function Distance_PointToPoint(pPoint1 As IJPoint, pPoint2 As IJPoint, pLine As IJLine) As Double
    Dim dDistance As Double: Let dDistance = 0
    
    Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromPoint(pPoint1)
    Dim pPosition2 As IJDPosition: Set pPosition2 = Position_FromPoint(pPoint2)
    
    Let dDistance = pPosition1.DistPt(pPosition2)
    If dDistance > EPSILON And Not pLine Is Nothing Then Let dDistance = dDistance * Sgn(Vector_FromLine(pLine).Dot(Vector_FromPositions(pPosition1, pPosition2)))
        
    ' return result
    Let Distance_PointToPoint = dDistance
End Function
Function Distance_PositionToPositionOrientedAlongVector(pPosition1 As IJDPosition, pPosition2 As IJDPosition, pVectorOfOrientation As IJDVector) As Double
    Dim pVector As IJDVector
    Set pVector = pPosition2.Subtract(pPosition1)
    
    Dim dDistance As Double
    Let dDistance = pVector.Length
    
    If Not pVectorOfOrientation Is Nothing Then
        If pVector.Dot(pVectorOfOrientation) < 0 Then Let dDistance = -dDistance
    End If
    
    ' return result
    Let Distance_PositionToPositionOrientedAlongVector = dDistance
End Function
Function Distance_PositionToLineOrientedAlongVector(pPosition As IJDPosition, pLine As IJLine, pVectorOfOrientation As IJDVector) As Double
    Dim pPositionProjected As IJDPosition
    Set pPositionProjected = Position_ProjectOnLine(pPosition, pLine)
    
     ' return result
    Let Distance_PositionToLineOrientedAlongVector = Distance_PositionToPositionOrientedAlongVector(pPositionProjected, pPosition, pVectorOfOrientation)
End Function
Function Length_FromCurve(pCurve As IJCurve) As Double
    Length_FromCurve = pCurve.Length
End Function
'
' processing Line
'
Sub Line_Debug(sMessage As String, pLine As IJLine)
    Dim dX0 As Double, dY0 As Double, dZ0 As Double
    Dim dX1 As Double, dY1 As Double, dZ1 As Double
    Call pLine.GetStartPoint(dX0, dY0, dZ0)
    Debug.Print (sMessage + ": X0= " + CStr(dX0) + ", Y0= " + CStr(dY0) + ". Z0= " + CStr(dZ0))
    Call pLine.GetEndPoint(dX1, dY1, dZ1)
    Debug.Print (sMessage + ": X1= " + CStr(dX1) + ", Y1= " + CStr(dY1) + ". Z1= " + CStr(dZ1))
End Sub
Function Position_FromLine(pLine As IJLine, iIndex As Integer) As IJDPosition
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
Function Position_FromLineString(pLineString As IJLineString, iIndex As Integer) As IJDPosition
    ' create new position
    Dim pPosition As IJDPosition
    Set pPosition = New DPosition
    
    ' retrieve coordinates
    Dim x As Double, y As Double, z As Double
    Call pLineString.GetPoint(iIndex, x, y, z)
    
    ' set coordinates
    Call pPosition.Set(x, y, z)
    
    ' return result
    Set Position_FromLineString = pPosition
End Function
Function Vector_FromLine(pLine As IJLine) As IJDVector
    ' get vector
    Dim u As Double, v As Double, w As Double
    Call pLine.GetDirection(u, v, w)
    
    ' create result
    Dim pVector As New DVector
    Call pVector.Set(u, v, w)
    
    ' return result
    Set Vector_FromLine = pVector
End Function
Public Function Line_InfiniteFromPositionAndVector(pPositionOfOrigin As IJDPosition, pVectorDirector As IJDVector) As IJLine
    ' create new line
    Dim dLength As Double
    Let dLength = BIG_EXTENSION
    
    Dim pVectorDirector1 As IJDVector
    Set pVectorDirector1 = pVectorDirector.Clone
    pVectorDirector1.Length = 1
    
    Dim pPositionOfInfiniteStartPoint As IJDPosition
    Set pPositionOfInfiniteStartPoint = pPositionOfOrigin.Offset(Vector_Scale(pVectorDirector1, -dLength))
    
    Dim pPositionOfInfiniteEndPoint As IJDPosition
    Set pPositionOfInfiniteEndPoint = pPositionOfOrigin.Offset(Vector_Scale(pVectorDirector1, dLength))
    
    Set Line_InfiniteFromPositionAndVector = GetGeometryFactory().Lines3d.CreateBy2Points(Nothing, _
        pPositionOfInfiniteStartPoint.x, pPositionOfInfiniteStartPoint.y, pPositionOfInfiniteStartPoint.z, _
        pPositionOfInfiniteEndPoint.x, pPositionOfInfiniteEndPoint.y, pPositionOfInfiniteEndPoint.z)
End Function
Public Function Line_InfiniteFromPositions(pPositionOfStartPoint As IJDPosition, pPositionOfEndPoint As IJDPosition) As IJLine
    ' create new line
    Dim dLength As Double
    Let dLength = BIG_EXTENSION
    
    Dim pVectorDirector As IJDVector
    Set pVectorDirector = pPositionOfEndPoint.Subtract(pPositionOfStartPoint)
    pVectorDirector.Length = 1
    
    Dim pPositionOfInfiniteStartPoint As IJDPosition
    Set pPositionOfInfiniteStartPoint = pPositionOfStartPoint.Offset(Vector_Scale(pVectorDirector, -dLength))
    
    Dim pPositionOfInfiniteEndPoint As IJDPosition
    Set pPositionOfInfiniteEndPoint = pPositionOfStartPoint.Offset(Vector_Scale(pVectorDirector, dLength))
    
    Set Line_InfiniteFromPositions = GetGeometryFactory().Lines3d.CreateBy2Points(Nothing, _
        pPositionOfInfiniteStartPoint.x, pPositionOfInfiniteStartPoint.y, pPositionOfInfiniteStartPoint.z, _
        pPositionOfInfiniteEndPoint.x, pPositionOfInfiniteEndPoint.y, pPositionOfInfiniteEndPoint.z)
End Function
Public Function Line_InfiniteFromLine(pLine As IJLine) As IJLine
    Dim pPositionOfOrigin As New DPosition
    Dim pVectorDirector As New DVector
    If True Then
        Dim dX As Double, dY As Double, dZ As Double
        Call pLine.GetStartPoint(dX, dY, dZ)
        Call pPositionOfOrigin.Set(dX, dY, dZ)
        
        Dim dU As Double, dV As Double, dW As Double
        Call pLine.GetDirection(dU, dV, dW)
        Call pVectorDirector.Set(dU, dV, dW)
    End If
    
    Set Line_InfiniteFromLine = Line_InfiniteFromPositionAndVector(pPositionOfOrigin, pVectorDirector)
End Function
Public Function Line_FromPositions(pPOM As IJDPOM, pPositionOfStartPoint As IJDPosition, pPositionOfEndPoint As IJDPosition) As IJLine
    ' create new line
    Set Line_FromPositions = GetGeometryFactory().Lines3d.CreateBy2Points(pPOM, _
        pPositionOfStartPoint.x, pPositionOfStartPoint.y, pPositionOfStartPoint.z, _
        pPositionOfEndPoint.x, pPositionOfEndPoint.y, pPositionOfEndPoint.z)
End Function
Sub Line_UpdateFromPositions(pLine As IJLine, pPosition0 As IJDPosition, pPosition1 As IJDPosition)
    Call pLine.DefineBy2Points(pPosition0.x, _
                               pPosition0.y, _
                               pPosition0.z, _
                               pPosition1.x, _
                               pPosition1.y, _
                               pPosition1.z)
End Sub
Function Position_ProjectOnLine(pPosition As IJDPosition, pLine As IJLine) As IJDPosition
    Dim pPositionOfStartPoint As IJDPosition
    Set pPositionOfStartPoint = Position_FromLine(pLine, 0)

    Dim pVectorOfReferencePoint As IJDVector
    Set pVectorOfReferencePoint = pPosition.Subtract(pPositionOfStartPoint)
    
    Dim pVectorOfLine As IJDVector
    Set pVectorOfLine = Vector_FromLine(pLine)
        
    Dim dProjectedDistanceFromStartPoint As Double
    Let dProjectedDistanceFromStartPoint = pVectorOfLine.Dot(pVectorOfReferencePoint)
    
    Dim pVectorOfProjectedPoint As New DVector
    Call pVectorOfProjectedPoint.Set(dProjectedDistanceFromStartPoint * pVectorOfLine.x, _
                                     dProjectedDistanceFromStartPoint * pVectorOfLine.y, _
                                     dProjectedDistanceFromStartPoint * pVectorOfLine.z)
                                     
    Set Position_ProjectOnLine = pPositionOfStartPoint.Offset(pVectorOfProjectedPoint)
End Function
Function AreLinesCoLinear(pLine1 As IJLine, pLine2 As IJLine) As Boolean
    Call DebugIn("AreLinesCoLinear")
    Let AreLinesCoLinear = False
    
    Dim pPositionStart1 As IJDPosition: Set pPositionStart1 = Position_FromLine(pLine1, 0)
    Dim pPositionEnd1 As IJDPosition: Set pPositionEnd1 = Position_FromLine(pLine1, 1)
    Dim pPositionStart2 As IJDPosition: Set pPositionStart2 = Position_FromLine(pLine2, 0)
    Dim pPositionEnd2 As IJDPosition: Set pPositionEnd2 = Position_FromLine(pLine2, 1)
    Call DebugValue("PositionStart1", pPositionStart1)
    Call DebugValue("PositionEnd1", pPositionEnd1)
    Call DebugValue("PositionStart2", pPositionStart2)
    Call DebugValue("PositionEnd2", pPositionEnd2)
    
    If IsPositionOnCurve(pPositionStart2, pLine1, True) _
    And IsPositionOnCurve(pPositionEnd2, pLine1, True) Then Let AreLinesCoLinear = True
    
    Call DebugOut
End Function
Function AreLinesCoPlanar(pLine1 As IJLine, pLine2 As IJLine) As Boolean
    Let AreLinesCoPlanar = True
    
    Dim pPositionStart1 As IJDPosition: Set pPositionStart1 = Position_FromLine(pLine1, 0)
    Dim pPositionStart2 As IJDPosition: Set pPositionStart2 = Position_FromLine(pLine2, 0)
    Dim pPositionEnd1 As IJDPosition: Set pPositionEnd1 = Position_FromLine(pLine1, 1)
    Dim pPositionEnd2 As IJDPosition: Set pPositionEnd2 = Position_FromLine(pLine2, 1)
    Dim pCurve1 As IJCurve: Set pCurve1 = Line_InfiniteFromPositions(pPositionStart1, pPositionEnd1)
    Dim pCurve2 As IJCurve: Set pCurve2 = Line_InfiniteFromPositions(pPositionStart2, pPositionEnd2)
    
    If pCurve1.IsPointOn(pPositionStart2.x, _
                         pPositionStart2.y, _
                         pPositionStart2.z) _
    Or pCurve1.IsPointOn(pPositionEnd2.x, _
                         pPositionEnd2.y, _
                         pPositionEnd2.z) Then Exit Function

    If pCurve2.IsPointOn(pPositionStart1.x, _
                         pPositionStart1.y, _
                         pPositionStart1.z) _
    Or pCurve2.IsPointOn(pPositionEnd1.x, _
                         pPositionEnd1.y, _
                         pPositionEnd1.z) Then Exit Function
    
    ' compute vector normal to the plane PointStart1, PointEnd1, PointStart2
    Dim pVectorNormal As IJDVector
    If True Then
        Dim pVectorOfLine1 As IJDVector:
        Set pVectorOfLine1 = pPositionEnd1.Subtract(pPositionStart1):
        Let pVectorOfLine1.Length = 1
        
        Dim pVectorBetweenStartPoints As IJDVector:
        Set pVectorBetweenStartPoints = pPositionStart2.Subtract(pPositionStart1):
        Let pVectorBetweenStartPoints.Length = 1
        
        Set pVectorNormal = pVectorOfLine1.Cross(pVectorBetweenStartPoints)
    End If
    
    ' test if vector of line2 is orthogonal to vector normal
    Dim pVectorOfLine2 As IJDVector: Set pVectorOfLine2 = pPositionEnd2.Subtract(pPositionStart2)
    If Abs(pVectorNormal.Dot(pVectorOfLine2)) > EPSILON Then
        Let AreLinesCoPlanar = False
    End If
End Function
Sub Line_Definition(pLine As IJLine, _
                        ByRef pPositionOfRoot As IJDPosition, _
                        ByRef pVectorOfDir As IJDVector, _
                        ByRef pPositionOfStart As IJDPosition, _
                        ByRef pPositionOfEnd As IJDPosition)
                                                                            
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    
    
    pLine.GetRootPoint dX, dY, dZ
    Set pPositionOfRoot = New DPosition
    pPositionOfRoot.Set dX, dY, dZ
    pLine.GetDirection dX, dY, dZ
    Set pVectorOfDir = New DVector
    pVectorOfDir.Set dX, dY, dZ
    pLine.GetStartPoint dX, dY, dZ
    Set pPositionOfStart = New DPosition
    pPositionOfStart.Set dX, dY, dZ
    pLine.GetEndPoint dX, dY, dZ
    Set pPositionOfEnd = New DPosition
    pPositionOfEnd.Set dX, dY, dZ

End Sub
Function Distance_LineToLine(pLine1 As IJLine, pLine2 As IJLine) As Double
    ' extract info from planes
    Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromLine(pLine1, 0)
    Dim pPosition2 As IJDPosition: Set pPosition2 = Position_FromLine(pLine2, 0)
    Dim pVector1 As IJDVector: Set pVector1 = Vector_FromLine(pLine1)
    Dim pVector2 As IJDVector: Set pVector2 = Vector_FromLine(pLine1)
                
    ' check if lines are parallel
    If True Then
        Dim pPositionOfVector1 As New DPosition: Call pPositionOfVector1.Set(pVector1.x, pVector1.y, pVector1.z)
        Dim pPositionOfVector2 As New DPosition: Call pPositionOfVector2.Set(pVector2.x, pVector2.y, pVector2.z)
        If pPositionOfVector1.DistPt(pPositionOfVector2) > EPSILON Then
            Err.Raise 1
        End If
    End If
        
    ' create infinite line
    Dim pLineInfinite As IJLine: Set pLineInfinite = Line_InfiniteFromPositionAndVector(pPosition1, pVector1)
        
    ' project position2 on infinite line
    Dim pPositionProjected2 As IJDPosition: Set pPositionProjected2 = Position_ProjectOnLine(pPosition2, pLineInfinite)
    
    ' return result
    Let Distance_LineToLine = pPositionProjected2.DistPt(pPosition2)
End Function
'
' processing Circle
'
Function Position_FromCircleCenter(pCircle As IJCircle) As IJDPosition
    Dim dX As Double, dY As Double, dZ As Double
    Call pCircle.GetCenterPoint(dX, dY, dZ)
    
    Dim pPosition As IJDPosition: Set pPosition = New DPosition
    Call pPosition.Set(dX, dY, dZ)
    
    Set Position_FromCircleCenter = pPosition
End Function
Function Vector_FromCircle(pCircle As IJCircle) As IJDVector
    Dim dU As Double, dV As Double, dW As Double
    Call pCircle.GetNormal(dU, dV, dW)
    
    Dim pVector As IJDVector: Set pVector = New DVector
    Call pVector.Set(dU, dV, dW)
    
    Set Vector_FromCircle = pVector
End Function
Public Function Circle_ByCenterNormalRadius(pPositionOfCenter As IJDPosition, _
                                             pVectorOfNormal As IJDVector, _
                                             dRadius As Double) As IJCircle
    Dim pCircle As New Circle3d
                                                                           
    pCircle.DefineByCenterNormalRadius pPositionOfCenter.x, pPositionOfCenter.y, pPositionOfCenter.z, _
                                       pVectorOfNormal.x, pVectorOfNormal.y, pVectorOfNormal.z, _
                                       dRadius
    Set Circle_ByCenterNormalRadius = pCircle
End Function
Sub Circle_Definition(pCircle As IJCircle, _
                            ByRef pPositionOfCenter As IJDPosition, _
                            ByRef pVectorOfNormal As IJDVector, _
                            ByRef dRadius As Double)
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    
    pCircle.GetCenterPoint dX, dY, dZ
    Set pPositionOfCenter = New DPosition
    pPositionOfCenter.Set dX, dY, dZ
    pCircle.GetNormal dX, dY, dZ
    Set pVectorOfNormal = New DVector
    pVectorOfNormal.Set dX, dY, dZ
    dRadius = pCircle.Radius

End Sub
Function Position_FromCircleStart(pCircle As IJCircle) As IJDPosition
    Dim dcx As Double
    Dim dcy As Double
    Dim dcz As Double
    Dim dnx As Double
    Dim dny As Double
    Dim dnz As Double
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    Dim dRadius As Double
    Dim dLength As Double
    Dim pPosition As DPosition
    
    pCircle.GetCenterPoint dcx, dcy, dcz
    pCircle.GetNormal dnx, dny, dnz
    dRadius = pCircle.Radius
    
    If Abs(dnx) > EPSILON Then
        dZ = 0
        dY = 1
        dX = -dny / dnx
    ElseIf Abs(dny) > EPSILON Then
            dX = 0
            dZ = 1
            dY = -dnz / dny
            ElseIf Abs(dnz) > EPSILON Then
                dY = 0
                dX = 1
                dZ = 0
                Else
                    Exit Function
            End If
            
    dLength = Sqr(dX * dX + dY * dY + dZ * dZ)
    Set pPosition = New DPosition
    pPosition.Set dcx + dRadius * dX / dLength, dcy + dRadius * dY / dLength, dcz + dRadius * dZ / dLength
    
    Set Position_FromCircleStart = pPosition
    
End Function
'
' processing Arc
'
Function Position_FromArcCenter(pArc As IJArc) As IJDPosition
    Dim dX As Double, dY As Double, dZ As Double
    Call pArc.GetCenterPoint(dX, dY, dZ)
    
    Dim pPosition As IJDPosition: Set pPosition = New DPosition
    Call pPosition.Set(dX, dY, dZ)
    
    Set Position_FromArcCenter = pPosition
End Function
Function Position_FromArc(pArc As IJArc, iIndex As Integer) As IJDPosition
    Dim dX As Double, dY As Double, dZ As Double
    If iIndex = 0 Then
        Call pArc.GetStartPoint(dX, dY, dZ)
    ElseIf iIndex = 1 Then
        Call pArc.GetEndPoint(dX, dY, dZ)
    Else
        Call pArc.GetCenterPoint(dX, dY, dZ)
    End If
    
    Dim pPosition As IJDPosition: Set pPosition = New DPosition
    Call pPosition.Set(dX, dY, dZ)
    
    Set Position_FromArc = pPosition
End Function
Function Vector_FromArc(pArc As IJArc) As IJDVector
    Dim dU As Double, dV As Double, dW As Double
    Call pArc.GetNormal(dU, dV, dW)
    
    Dim pVector As IJDVector: Set pVector = New DVector
    Call pVector.Set(dU, dV, dW)
    
    Set Vector_FromArc = pVector
End Function
Sub Arc_Definition(pArc As IJArc, _
                        ByRef pPositionOfCenter As IJDPosition, _
                        ByRef pVectorOfNormal As IJDVector, _
                        ByRef dRadius As Double, _
                        ByRef pPositionOfStart As IJDPosition, _
                        ByRef pPositionOfEnd As IJDPosition)
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    
    
    pArc.GetCenterPoint dX, dY, dZ
    Set pPositionOfCenter = New DPosition
    pPositionOfCenter.Set dX, dY, dZ
    pArc.GetNormal dX, dY, dZ
    Set pVectorOfNormal = New DVector
    pVectorOfNormal.Set dX, dY, dZ
    dRadius = pArc.Radius
    pArc.GetStartPoint dX, dY, dZ
    Set pPositionOfStart = New DPosition
    pPositionOfStart.Set dX, dY, dZ
    pArc.GetEndPoint dX, dY, dZ
    Set pPositionOfEnd = New DPosition
    pPositionOfEnd.Set dX, dY, dZ
End Sub
Function Vector_TangentFromArc(pArc As IJArc, _
                               pPosition As IJDPosition) As IJDVector
    If Not IsPositionOnCurve(pPosition, pArc) Then
        Exit Function
    End If
    
    Dim pPositionOfCenter As New DPosition
    Dim pVectorOfNormal As New DVector
    Dim pVectorOfTangent As DVector
    Dim pVectorOfRadius As IJDVector
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    
    pArc.GetCenterPoint dX, dY, dZ
    pPositionOfCenter.Set dX, dY, dZ
    pArc.GetNormal dX, dY, dZ
    pVectorOfNormal.Set dX, dY, dZ
    
    Set pVectorOfRadius = pPosition.Subtract(pPositionOfCenter)
    Set pVectorOfTangent = pVectorOfNormal.Cross(pVectorOfRadius)
    
    Set Vector_TangentFromArc = Vector_Normalize(pVectorOfTangent)
End Function
Public Function Arc_Concentric(pArc As IJArc, _
                               dOffsetValue As Double) As IJArc
                                         
    Dim pArcConcentric As Arc3d
    Dim pPositionOfCenter As IJDPosition
    Dim pVectorOfNormal As IJDVector
    Dim dRadius As Double
    Dim pPositionOfStart As IJDPosition
    Dim pPositionOfEnd As IJDPosition
    Dim pPositionNewOfStart As DPosition
    Dim pPositionNewOfEnd As DPosition
    Dim pVectorOfRadiusStart As IJDVector
    Dim pVectorOfRadiusEnd As IJDVector
    Dim dRatio As Double
    
    Arc_Definition pArc, pPositionOfCenter, pVectorOfNormal, dRadius, pPositionOfStart, pPositionOfEnd
    
    dRatio = (dOffsetValue + dRadius) / dRadius
    Set pVectorOfRadiusStart = Vector_FromPositions(pPositionOfCenter, pPositionOfStart)
    Set pVectorOfRadiusEnd = Vector_FromPositions(pPositionOfCenter, pPositionOfEnd)
    
    Set pPositionNewOfStart = New DPosition
    Set pPositionNewOfStart = pPositionOfCenter.Offset(Vector_Scale(pVectorOfRadiusStart, dRatio))
    
    Set pPositionNewOfEnd = New DPosition
    Set pPositionNewOfEnd = pPositionOfCenter.Offset(Vector_Scale(pVectorOfRadiusEnd, dRatio))
  
    Set pArcConcentric = New Arc3d
    Arc_ModifyByCenterNormalStartEnd pArcConcentric, pPositionOfCenter, pVectorOfNormal, pPositionNewOfStart, pPositionNewOfEnd
  
    Set Arc_Concentric = pArcConcentric
    
End Function
Public Function Arc_ModifyByCenterNormalStartEnd(pArc As IJArc, _
                                            pPositionOfCenter As IJDPosition, _
                                            pVectorOfNormal As IJDVector, _
                                            pPositionOfStart As IJDPosition, _
                                            pPositionOfEnd As IJDPosition)
    pArc.DefineByCtrNormStartEnd pPositionOfCenter.x, pPositionOfCenter.y, pPositionOfCenter.z, _
                                 pVectorOfNormal.x, pVectorOfNormal.y, pVectorOfNormal.z, _
                                 pPositionOfStart.x, pPositionOfStart.y, pPositionOfStart.z, _
                                 pPositionOfEnd.x, pPositionOfEnd.y, pPositionOfEnd.z
End Function
'
' processing Curve
'
Function Position_FromCurve(pCurve As IJCurve, iIndex As Integer) As IJDPosition
    ' create new position
    Dim pPosition As IJDPosition
    Set pPosition = New DPosition
    
    ' retrieve coordinates
    Dim x0 As Double, y0 As Double, z0 As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    Call pCurve.EndPoints(x0, y0, z0, x1, y1, z1)
       
    ' set coordinates
    If iIndex = 0 Then
        Call pPosition.Set(x0, y0, z0)
    ElseIf iIndex = 1 Then
        Call pPosition.Set(x1, y1, z1)
    End If
    
    ' return result
    Set Position_FromCurve = pPosition
End Function
Public Function Vector_NormalFromCurve(pCurve As IJCurve) As IJDVector
    Set Vector_NormalFromCurve = Nothing
    
    Dim eGeom3dSurfaceScopeConstants As Geom3dSurfaceScopeConstants
    Dim dU As Double, dV As Double, dW As Double
    Call pCurve.Normal(eGeom3dSurfaceScopeConstants, dU, dV, dW)
    If eGeom3dSurfaceScopeConstants = SURFACE_SCOPE_PLANAR Then
        Dim pVectorOfNormal As New DVector
        Call pVectorOfNormal.Set(dU, dV, dW)
        Set Vector_NormalFromCurve = pVectorOfNormal
    End If
End Function
Public Function Position_AtCurvesIntersection(pCurve1 As IJCurve, pCurve2 As IJCurve, pPointOfTrackPoint As IJPoint, ByVal iTrackFlag As Integer) As IJDPosition
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
    
    ' choose the nearer/further intersection position
    Dim pPositionOfResultPoint As IJDPosition:
    Dim pPositionOfTrackPoint As IJDPosition: Set pPositionOfTrackPoint = Nothing
    If Not pPointOfTrackPoint Is Nothing Then Set pPositionOfTrackPoint = Position_FromPoint(pPointOfTrackPoint)
    Set pPositionOfResultPoint = Position_FromPositionsV1(pElementsOfPositions, pPositionOfTrackPoint, iTrackFlag)
    
    ' return result
    Set Position_AtCurvesIntersection = pPositionOfResultPoint
End Function
Public Function Positions_AtCurvesIntersection(pCurve1 As IJCurve, pCurve2 As IJCurve) As IJElements
    ' initialize result
    Set Positions_AtCurvesIntersection = Nothing
    
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
    End If
    
    ' return result
    Set Positions_AtCurvesIntersection = pElementsOfPositions
End Function
Public Function Position_AtCurveSurfaceIntersection(pCurve1 As IJCurve, pSurface2 As IJSurface, _
                                                    pPointOfTrackPoint As IJPoint, ByVal iTrackFlag As Integer) As IJDPosition
    ' initialize result
    Set Position_AtCurveSurfaceIntersection = Nothing
    
    ' compute all possible positions
    Dim pElementsOfPositions As IJElements
    If True Then
        ' compute positions as an array of doubles
        Dim dDoubles() As Double
        Dim lNumIntersections As Long: Let lNumIntersections = 0
        If True Then
            Dim lNumOverlaps As Long
            Dim eGeom3dIntersectConstants As Geom3dIntersectConstants
            On Error Resume Next
            Call pCurve1.Intersect(pSurface2, lNumIntersections, dDoubles, lNumOverlaps, eGeom3dIntersectConstants)
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
    
    ' choose the nearer/further intersection position
    Dim pPositionOfResultPoint As IJDPosition:
    Dim pPositionOfTrackPoint As IJDPosition: Set pPositionOfTrackPoint = Nothing
    If Not pPointOfTrackPoint Is Nothing Then Set pPositionOfTrackPoint = Position_FromPoint(pPointOfTrackPoint)
    Set pPositionOfResultPoint = Position_FromPositionsV1(pElementsOfPositions, pPositionOfTrackPoint, iTrackFlag)
    
    ' return result
    Set Position_AtCurveSurfaceIntersection = pPositionOfResultPoint
End Function
Public Function IsPositionOnCurve(pPosition As IJDPosition, pCurve As IJCurve, Optional bExtendCurve As Boolean = False) As Boolean
    Dim pCurveTemp As IJCurve
    If TypeOf pCurve Is IJLine And bExtendCurve Then
        Set pCurveTemp = Line_InfiniteFromLine(pCurve)
    Else
        Set pCurveTemp = pCurve
    End If
    Let IsPositionOnCurve = pCurveTemp.IsPointOn(pPosition.x, pPosition.y, pPosition.z)
End Function
Public Function IsPositionInsideCurve(pPosition As IJDPosition, pCurve As IJCurve) As Boolean
    Let IsPositionInsideCurve = False
    If pCurve.IsPointOn(pPosition.x, pPosition.y, pPosition.z) Then
        Dim pStartPoint As New DPosition
        Dim pEndPoint As New DPosition
        If True Then
            Dim dX0 As Double, dY0 As Double, dZ0 As Double, dX1 As Double, dY1 As Double, dZ1 As Double
            Call pCurve.EndPoints(dX0, dY0, dZ0, dX1, dY1, dZ1)
            Call pStartPoint.Set(dX0, dY0, dZ0)
            Call pEndPoint.Set(dX1, dY1, dZ1)
        End If
        If pPosition.DistPt(pStartPoint) > EPSILON _
        And pPosition.DistPt(pEndPoint) > EPSILON Then
            Let IsPositionInsideCurve = True
        End If
    End If
End Function
Public Function IsPositionOnSurface(pPosition As IJDPosition, pSurf As IJSurface) As Boolean
    Let IsPositionOnSurface = pSurf.IsPointOn(pPosition.x, pPosition.y, pPosition.z)
End Function
Public Function AreCurvesInContact(pCurve1 As IJCurve, pCurve2 As IJCurve) As Boolean
    Let AreCurvesInContact = True
    
    Dim pPositionStart1 As IJDPosition: Set pPositionStart1 = Position_FromCurve(pCurve1, 0)
    Dim pPositionStart2 As IJDPosition: Set pPositionStart2 = Position_FromCurve(pCurve2, 0)
    Dim pPositionEnd1 As IJDPosition: Set pPositionEnd1 = Position_FromCurve(pCurve1, 1)
    Dim pPositionEnd2 As IJDPosition: Set pPositionEnd2 = Position_FromCurve(pCurve2, 1)
    
    If pCurve1.IsPointOn(pPositionStart2.x, _
                         pPositionStart2.y, _
                         pPositionStart2.z) _
    Or pCurve1.IsPointOn(pPositionEnd2.x, _
                         pPositionEnd2.y, _
                         pPositionEnd2.z) Then Exit Function

    If pCurve2.IsPointOn(pPositionStart1.x, _
                         pPositionStart1.y, _
                         pPositionStart1.z) _
    Or pCurve2.IsPointOn(pPositionEnd1.x, _
                         pPositionEnd1.y, _
                         pPositionEnd1.z) Then Exit Function
                         
    Let AreCurvesInContact = False
End Function
Public Function AreCurvesEndPointConnected(pCurve1 As IJCurve, pCurve2 As IJCurve) As Boolean
    Let AreCurvesEndPointConnected = False
    
    Dim pPositionStart1 As IJDPosition: Set pPositionStart1 = Position_FromCurve(pCurve1, 0)
    Dim pPositionStart2 As IJDPosition: Set pPositionStart2 = Position_FromCurve(pCurve2, 0)
    Dim pPositionEnd1 As IJDPosition: Set pPositionEnd1 = Position_FromCurve(pCurve1, 1)
    Dim pPositionEnd2 As IJDPosition: Set pPositionEnd2 = Position_FromCurve(pCurve2, 1)
    
    If pPositionStart1.DistPt(pPositionStart2) < EPSILON _
    Or pPositionStart1.DistPt(pPositionEnd2) < EPSILON _
    Or pPositionEnd1.DistPt(pPositionStart2) < EPSILON _
    Or pPositionEnd1.DistPt(pPositionEnd2) < EPSILON Then Let AreCurvesEndPointConnected = True
End Function
'
' processing Plane
'
Public Function Position_FromPlane(pPlane As IJPlane) As IJDPosition
    Dim dX As Double, dY As Double, dZ As Double
    Call pPlane.GetRootPoint(dX, dY, dZ)

    Dim pPosition As New DPosition
    Call pPosition.Set(dX, dY, dZ)

    Set Position_FromPlane = pPosition
End Function
Public Function Vector_FromPlane(pPlane As IJPlane) As IJDVector
    Dim dU As Double, dV As Double, dW As Double
    Call pPlane.GetNormal(dU, dV, dW)

    Dim pVector As New DVector
    Call pVector.Set(dU, dV, dW)

    Set Vector_FromPlane = pVector
End Function
Public Function Position_ProjectOnPlane(pPosition As IJDPosition, pPlane As IJPlane) As IJDPosition
    ' retrieve root point and normal of transient plane
    Dim pPositionOfRootPoint As IJDPosition
    Dim pVectorOfNormal As IJDVector
    If True Then
        Set pPositionOfRootPoint = Position_FromPlane(pPlane)
        Set pVectorOfNormal = Vector_FromPlane(pPlane)
    End If

    ' project cursor to sketching plane
    Dim pVector As IJDVector
    Set pVector = pPositionOfRootPoint.Subtract(pPosition)

    Dim dDistanceNormal As Double
    Let dDistanceNormal = pVectorOfNormal.Dot(pVector)

    Dim pPositionProjected As IJDPosition
    Set pPositionProjected = pPosition.Offset(Vector_Scale(pVectorOfNormal, dDistanceNormal))

    ' return result
    Set Position_ProjectOnPlane = pPositionProjected
End Function
Public Function Vector_FromPlanarCurve(pCurve As IJCurve) As IJDVector
    ' prepare result
    Dim pVector As IJDVector: Set pVector = Nothing
        
    If pCurve.Scope = CURVE_SCOPE_PLANAR Then
        Dim dU As Double, dV As Double, dW As Double
        Call pCurve.Normal(CURVE_SCOPE_PLANAR, dU, dV, dW)
        Set pVector = Vector_FromCoordinates(dU, dV, dW)
    End If
    
    ' return result
    Set Vector_FromPlanarCurve = pVector
End Function

Public Function Plane_ByPositionAndNormal(pPositionOfRootPoint As IJDPosition, pVectorOfNormal As IJDVector) As IJPlane
    Set Plane_ByPositionAndNormal = GetGeometryFactory().Planes3d.CreateByPointNormal(Nothing, _
                                        pPositionOfRootPoint.x, _
                                        pPositionOfRootPoint.y, _
                                        pPositionOfRootPoint.z, _
                                        pVectorOfNormal.x, _
                                        pVectorOfNormal.y, _
                                        pVectorOfNormal.z)
End Function
Public Sub Plane_Definition(pPlane As IJPlane, _
                                 ByRef pPositionOfRootPoint As IJDPosition, _
                                 ByRef pVectorOfNormal As IJDVector)
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    
    pPlane.GetRootPoint dX, dY, dZ
    Set pPositionOfRootPoint = New DPosition
    pPositionOfRootPoint.Set dX, dY, dZ
    pPlane.GetNormal dX, dY, dZ
    Set pVectorOfNormal = New DVector
    pVectorOfNormal.Set dX, dY, dZ
End Sub
Function Distance_PlaneToPlane(pPlane1 As IJPlane, pPlane2 As IJPlane) As Double
    ' extract info from planes
    Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromPlane(pPlane1)
    Dim pPosition2 As IJDPosition: Set pPosition2 = Position_FromPlane(pPlane2)
    Dim pVector1 As IJDVector: Set pVector1 = Vector_FromPlane(pPlane1)
    Dim pVector2 As IJDVector: Set pVector2 = Vector_FromPlane(pPlane2)
                
    ' check if planes are parallel
    If True Then
        Dim pPositionOfVector1 As New DPosition: Call pPositionOfVector1.Set(pVector1.x, pVector1.y, pVector1.z)
        Dim pPositionOfVector2 As New DPosition: Call pPositionOfVector2.Set(pVector2.x, pVector2.y, pVector2.z)
        If pPositionOfVector1.DistPt(pPositionOfVector2) > EPSILON _
        And pPositionOfVector1.DistPt(pPositionOfVector2) < 2 - EPSILON Then
            Err.Raise 1
        End If
    End If
        
    ' create infinite plane
    Dim pPlaneInfinite As IJPlane: Set pPlaneInfinite = Plane_ByPositionAndNormal(pPosition1, pVector1)
        
    ' project position2 on infinite plane
    Dim pPositionProjected2 As IJDPosition: Set pPositionProjected2 = Position_ProjectOnPlane(pPosition2, pPlaneInfinite)
        
    ' return result
    Let Distance_PlaneToPlane = pPosition2.DistPt(pPositionProjected2)
End Function
Function Distance_PointToPlane(pPoint1 As IJPoint, pPlane2 As IJPlane) As Double
    Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromPoint(pPoint1)
    Dim pPosition2 As IJDPosition: Set pPosition2 = Position_FromPlane(pPlane2)
    Dim pVector2 As IJDVector: Set pVector2 = Vector_FromPlane(pPlane2)

    ' create infinite plane
    Dim pPlaneInfinite As IJPlane: Set pPlaneInfinite = Plane_ByPositionAndNormal(pPosition2, pVector2)

    ' project position2 on infinite plane
    Dim pPositionProjected1 As IJDPosition: Set pPositionProjected1 = Position_ProjectOnPlane(pPosition1, pPlaneInfinite)
        
    ' return result
    Let Distance_PointToPlane = pPosition1.DistPt(pPositionProjected1)
End Function
'
' processing Surface
'
Public Function IsLineOnSurface(pSurface As IJSurface, pLine As IJLine) As Boolean
    Let IsLineOnSurface = False
    
    Dim pPositionStart As IJDPosition: Set pPositionStart = Position_FromLine(pLine, 0)
    Dim pPositionEnd As IJDPosition: Set pPositionEnd = Position_FromLine(pLine, 1)
    If pSurface.IsPointOn(pPositionStart.x, _
                          pPositionStart.y, _
                          pPositionStart.z) _
    And pSurface.IsPointOn(pPositionEnd.x, _
                           pPositionEnd.y, _
                           pPositionEnd.z) Then Let IsLineOnSurface = True
End Function
Public Function Position_FromSurface(pSurface As IJSurface, dU As Double, dV As Double) As IJDPosition
    ' get center point
    Dim x As Double, y As Double, z As Double
    Call pSurface.Position(dU, dV, x, y, z)

    ' store it
    Dim pPosition As IJDPosition: Set pPosition = New DPosition:
    Call pPosition.Set(x, y, z)

    ' return it
    Set Position_FromSurface = pPosition
End Function
Public Function Angle_GetFromCenterStartEndAndNormal(pPositionOfCenterPoint As IJDPosition, _
                                                     pPositionOfStartPoint As IJDPosition, _
                                                     pPositionOfEndPoint As IJDPosition, _
                                                     pVectorOfNormal As IJDVector) As Double
'''        Call DebugIn("Angle_GetFromCenterStartEndAndNormal")
'''        Call DebugInput("PositionOfCenterPoint", pPositionOfCenterPoint)
'''        Call DebugInput("PositionOfStartPoint", pPositionOfStartPoint)
'''        Call DebugInput("PositionOfEndPoint", pPositionOfEndPoint)
'''        Call DebugInput("VectorOfNormal", pVectorOfNormal)
        
         ' initialize result
        Dim dAngle As Double: Let dAngle = 0
        
        If pPositionOfStartPoint.DistPt(pPositionOfEndPoint) > EPSILON Then
            ' compute arc
            Dim pArc As IJArc
            Set pArc = GetGeometryFactory().Arcs3d.CreateByCtrNormStartEnd(Nothing, _
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
            ' extract angle
            Let dAngle = pArc.SweepAngle ' / PI * 180
        End If
        
        ' return result
        Let Angle_GetFromCenterStartEndAndNormal = dAngle

'''        Call DebugOutput("Angle", Angle_GetFromCenterStartEndAndNormal)
'''        Call DebugOut
End Function
Public Function Lines_GetMinimumLength(pLines() As IJLine) As Double
    ' prepare result
    Dim dMinimumLength As Double
    Let dMinimumLength = 1000000#
    
    ' retrieve size of array
    Dim iSize As Integer: Let iSize = UBound(pLines, 1)
    
    ' loop on lines
    Dim i As Integer
    For i = 1 To iSize
        Dim dLength As Double: Let dLength = pLines(i).Length()
        If dLength < dMinimumLength Then Let dMinimumLength = dLength
    Next

    ' return result
    Let Lines_GetMinimumLength = dMinimumLength
End Function
Function GetLineAtNormalToLines(pGCGeomFactory As IJGCGeomFactory, pGCGeomFactory2 As IJGCGeomFactory2, pPOM As IJDPOM, pLines() As IJLine) As IJLine
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
Function GetLineAtNormalToLinesAndParallelToLine(pGCGeomFactory As IJGCGeomFactory, pGCGeomFactory2 As IJGCGeomFactory2, pPOM As IJDPOM, pLines() As IJLine, pLineOfOrthogonalLine As IJLine) As IJLine
    ' prepare result
    Dim pLineOfNormal As IJLine
    
    ' retrieve size of array
    Dim iLines As Integer: Let iLines = UBound(pLines, 1)
    
    ' verify that all lines are in the same plane
    Dim bAreLinesCoplanar As Boolean: Let bAreLinesCoplanar = True
    If True Then
        Dim pVectorOfNormal As IJDVector: Set pVectorOfNormal = Vector_FromLine(pLineOfOrthogonalLine)
        ' loop on lines
        Dim i As Integer
        For i = 1 To iLines
            If Not AreVectorsPerpendicular(pVectorOfNormal, Vector_FromLine(pLines(i))) Then
                Let bAreLinesCoplanar = False
                Exit For
            End If
        Next
    End If
    
    ' recompute normal line
    If bAreLinesCoplanar Then
        If iLines > 1 Then
            'take into account the selection order of the regular members
            Set pLineOfNormal = GetLineAtNormalToLines(pGCGeomFactory, pGCGeomFactory, pPOM, pLines)
        End If
        
        ' if iLines = 1 or the previous compute has failed
        If pLineOfNormal Is Nothing Then
            ' orientate the normal on the side where the orthogonal line is the longer
            Dim pCoordinateSystem As IJDCoordinateSystem: Set pCoordinateSystem = pGCGeomFactory2.CSByLines(pPOM, pLineOfOrthogonalLine, pLines(1), Nothing, GCZX, GCDirect)
            Set pLineOfNormal = pGCGeomFactory.LineFromCS.PlaceGeometry(pPOM, pCoordinateSystem, GCZAxis, 1)
        End If
    End If
    
    ' return result
    Set GetLineAtNormalToLinesAndParallelToLine = pLineOfNormal
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

    ' order array of angles
    Call Angles_OrderByAngle(dAngles, dOrderedAngles, lOrderedIndexes)
End Sub
Function ArePlanesCoplanar(pPlane1 As IJPlane, pPlane2 As IJPlane) As Boolean
    Let ArePlanesCoplanar = False
    Dim pVector1 As IJDVector: Set pVector1 = Vector_FromPlane(pPlane1)
    Dim pVector2 As IJDVector: Set pVector2 = Vector_FromPlane(pPlane2)
    If Abs(pVector1.Dot(pVector2)) > 0.999 Then
'        Call ShowMsg("Vectors are parallel")
        Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromPlane(pPlane1)
        Dim pPosition2 As IJDPosition: Set pPosition2 = Position_FromPlane(pPlane2)
        Dim pVector As IJDVector: Set pVector = pPosition2.Subtract(pPosition1)
        'Let pVector.Length = 1
        Dim dNonCoplanearity As Double
        Let dNonCoplanearity = Abs(pVector.Dot(pVector1))
'        Call ShowMsg("NonCoplanearity= " + CStr(dNonCoplanearity))
        If Abs(dNonCoplanearity) < 0.001 Then
'            Call ShowMsg("Planes are coplanar")
            Let ArePlanesCoplanar = True
        End If
    End If
End Function
Function ArePlanesParallel(pPlane1 As IJPlane, pPlane2 As IJPlane) As Boolean
    Let ArePlanesParallel = False
    Dim pVector1 As IJDVector: Set pVector1 = Vector_FromPlane(pPlane1)
    Dim pVector2 As IJDVector: Set pVector2 = Vector_FromPlane(pPlane2)
    If Abs(pVector1.Dot(pVector2)) > 0.999 Then
        Let ArePlanesParallel = True
    End If
End Function
Function GetPositionAtExtremityOfLineOnCurve(pLine As IJLine, pCurve As IJCurve) As IJDPosition
    ' initialize result
    Dim pPositionAtExtremity As IJDPosition
    
    ' retrieve line extremities
    Dim pPosition0 As IJDPosition: Set pPosition0 = Position_FromLine(pLine, 0)
    Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromLine(pLine, 1)
    
    ' compute result
    If IsPositionOnCurve(pPosition0, pCurve) Then
        Set pPositionAtExtremity = pPosition0
    ElseIf IsPositionOnCurve(pPosition1, pCurve) Then
        Set pPositionAtExtremity = pPosition1
    End If
    
    ' return result
    Set GetPositionAtExtremityOfLineOnCurve = pPositionAtExtremity
End Function
Function GetPositionAtCommonExtremityOfLines(pLines() As IJLine) As IJDPosition
    ' initialize result
    Dim pPositionOfCommonExtremity As IJDPosition
    
    ' retrieve size of array
    Dim iSize As Integer: Let iSize = UBound(pLines, 1)
    
    ' only process if only 2 lines
    If iSize >= 2 Then
        ' retrieve extremities of the 2 first lines
        Dim pPosition10 As IJDPosition: Set pPosition10 = Position_FromLine(pLines(1), 0)
        Dim pPosition11 As IJDPosition: Set pPosition11 = Position_FromLine(pLines(1), 1)
        Dim pPosition20 As IJDPosition: Set pPosition20 = Position_FromLine(pLines(2), 0)
        Dim pPosition21 As IJDPosition: Set pPosition21 = Position_FromLine(pLines(2), 1)
        
        ' find if they share a common extremity
        If pPosition10.DistPt(pPosition20) < EPSILON Then
            Set pPositionOfCommonExtremity = pPosition10
        ElseIf pPosition10.DistPt(pPosition21) < EPSILON Then
            Set pPositionOfCommonExtremity = pPosition10
        ElseIf pPosition11.DistPt(pPosition20) < EPSILON Then
            Set pPositionOfCommonExtremity = pPosition11
        ElseIf pPosition11.DistPt(pPosition21) < EPSILON Then
            Set pPositionOfCommonExtremity = pPosition11
        End If
        
        If iSize > 2 And Not pPositionOfCommonExtremity Is Nothing Then
            ' find if other lines also share the same common extremity
            Dim i As Integer
            For i = 3 To iSize
                If pPositionOfCommonExtremity.DistPt(Position_FromLine(pLines(i), 0)) > EPSILON _
                And pPositionOfCommonExtremity.DistPt(Position_FromLine(pLines(i), 1)) > EPSILON Then
                    Set pPositionOfCommonExtremity = Nothing
                    Exit For
                End If
            Next
        End If
    End If
    
    ' return result
    Set GetPositionAtCommonExtremityOfLines = pPositionOfCommonExtremity
End Function
Private Function GetCSByLines(pGCGeomFactory As IJGCGeomFactory, pGCGeomFactory2 As IJGCGeomFactory2, pPOM As IJDPOM, pLines() As IJLine) As IJDCoordinateSystem
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
'
' private functions
'
Private Function GetPositionsOfExtremitiesOfLines(pLines() As IJLine, pPositionOfNode As IJDPosition) As IJDPosition()
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
Private Function GetAnglesFromCenterExtremitiesAndNormal(pPositionOfCenter As IJDPosition, _
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
Public Function GetAngleFrom2LinesAndNormal(pLine1 As IJLine, pLine2 As IJLine, pLineOfNormal As IJLine) As Double
    Let GetAngleFrom2LinesAndNormal = GetAngleFromCenterStartEndAndNormal(Position_FromLine(pLine1, 0), _
                                                                          Position_FromLine(pLine1, 1), _
                                                                          Position_FromLine(pLine2, 1), _
                                                                          Vector_FromLine(pLineOfNormal))
End Function
Public Function GetAngleFromCenterStartEndAndNormal(pPositionOfCenterPoint As IJDPosition, _
                                             pPositionOfStartPoint As IJDPosition, _
                                             pPositionOfEndPoint As IJDPosition, _
                                             pVectorOfNormal As IJDVector) As Double
        Call DebugIn("GetAngleFromCenterStartEndAndNormal")
        Call DebugInput("PositionOfCenterPoint", pPositionOfCenterPoint)
        Call DebugInput("PositionOfStartPoint", pPositionOfStartPoint)
        Call DebugInput("PositionOfEndPoint", pPositionOfEndPoint)
        Call DebugInput("VectorOfNormal", pVectorOfNormal)
        
        ' compute arc
        Dim pArc As IJArc
        Set pArc = GetGeometryFactory().Arcs3d.CreateByCtrNormStartEnd(Nothing, _
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
        
        Call DebugOutput("Angle", GetAngleFromCenterStartEndAndNormal)
        Call DebugOut
End Function
Private Function GetElementsByOrderingByAngle(pElements As IJElements, dAngles() As Double) As IJElements
    ' retrieve size of collection and array
    Dim iCount As Integer: Let iCount = pElements.Count
    
    ' remember already processed elements
    Dim bIsAlreadyProcessed() As Boolean: ReDim bIsAlreadyProcessed(1 To iCount)
    Dim i As Integer: For i = 1 To iCount: Let bIsAlreadyProcessed(i) = False: Next
    
    ' initialize the ordered collection
    Dim pElementsOfOrderedElements As IJElements: Set pElementsOfOrderedElements = New JObjectCollection
    Call pElementsOfOrderedElements.Add(pElements.Item(1))
    Let bIsAlreadyProcessed(1) = True
    
    ' fill the ordered collection
    For i = 2 To iCount
        Dim dAngleMini As Double: Let dAngleMini = 360 ' degrees
        Dim j As Integer: Dim jOfSmallerAngle As Integer
        For j = 2 To iCount
            If bIsAlreadyProcessed(j) = False Then
                If dAngles(j) < dAngleMini Then
                    Let dAngleMini = dAngles(j)
                    Let jOfSmallerAngle = j
                End If
            End If
        Next
        Call pElementsOfOrderedElements.Add(pElements.Item(jOfSmallerAngle))
        Let bIsAlreadyProcessed(jOfSmallerAngle) = True
    Next

    ' return result
    Set GetElementsByOrderingByAngle = pElementsOfOrderedElements
End Function
Private Sub Angles_OrderByAngle(dAngles() As Double, ByRef dOrderedAngles() As Double, ByRef lOrderedIndexes() As Long)
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
Function Position_CommonEndPointFromLines(pLines() As IJLine) As IJDPosition
    ' initialize result
    Dim pPositionOfCommonExtremity As IJDPosition
    
    ' retrieve size of array
    Dim iSize As Integer: Let iSize = UBound(pLines, 1)
    
    ' only process if only 2 lines
    If iSize >= 2 Then
        ' retrieve extremities of the 2 first lines
        Dim pPosition10 As IJDPosition: Set pPosition10 = Position_FromLine(pLines(1), 0)
        Dim pPosition11 As IJDPosition: Set pPosition11 = Position_FromLine(pLines(1), 1)
        Dim pPosition20 As IJDPosition: Set pPosition20 = Position_FromLine(pLines(2), 0)
        Dim pPosition21 As IJDPosition: Set pPosition21 = Position_FromLine(pLines(2), 1)
        
        ' find if they share a common extremity
        If pPosition10.DistPt(pPosition20) < EPSILON Then
            Set pPositionOfCommonExtremity = pPosition10
        ElseIf pPosition10.DistPt(pPosition21) < EPSILON Then
            Set pPositionOfCommonExtremity = pPosition10
        ElseIf pPosition11.DistPt(pPosition20) < EPSILON Then
            Set pPositionOfCommonExtremity = pPosition11
        ElseIf pPosition11.DistPt(pPosition21) < EPSILON Then
            Set pPositionOfCommonExtremity = pPosition11
        End If
        
        If iSize > 2 And Not pPositionOfCommonExtremity Is Nothing Then
            ' find if other lines also share the same common extremity
            Dim i As Integer
            For i = 3 To iSize
                If pPositionOfCommonExtremity.DistPt(Position_FromLine(pLines(i), 0)) > EPSILON _
                And pPositionOfCommonExtremity.DistPt(Position_FromLine(pLines(i), 1)) > EPSILON Then
                    Set pPositionOfCommonExtremity = Nothing
                    Exit For
                End If
            Next
        End If
    End If
    
    ' return result
    Set Position_CommonEndPointFromLines = pPositionOfCommonExtremity
End Function


