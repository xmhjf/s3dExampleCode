Attribute VB_Name = "AutoMathHelpers"
Option Explicit
' depends on AutoMaths (IJDPosition, IJDVector, ...)
' depends on SP3DPositionAndOrientation (IJLocalCoordinatesystem)
' depends on CoreCollections (IJElements)

'
' processing Position
'
Sub Position_Debug(pPosition As IJDPosition, sText As String)
    Debug.Print sText, pPosition.x, pPosition.y, pPosition.z
End Sub
Function Position_ToString(pPosition As IJDPosition, Optional sText As String = "") As String
    Let Position_ToString = ""
    If Not sText = "" Then Let Position_ToString = sText + ": "
    Let Position_ToString = Position_ToString + "x= " + CStr(Round(pPosition.x, 6)) + ", y= " + CStr(Round(pPosition.y, 6)) + ", z= " + CStr(Round(pPosition.z, 6))
End Function
Public Function Position_FromPosition(pPosition As IJDPosition) As IJDPosition
    Set Position_FromPosition = pPosition
End Function
Function Position_FromPositionsV1(pElementsOfPositions As IJElements, pPositionOfTrackPoint As IJDPosition, ByVal iTrackFlag As Integer) As IJDPosition
    ' initialize result
    Set Position_FromPositionsV1 = Nothing
    
    ' choose the nearer/further intersection point
    Dim pPositionOfResultPoint As IJDPosition
    ' V0 : If pPointOfTrackPoint Is Nothing Then
    If pPositionOfTrackPoint Is Nothing Then
        Set pPositionOfResultPoint = pElementsOfPositions(1)
        If iTrackFlag = FAR And pElementsOfPositions.Count > 1 Then Set pPositionOfResultPoint = pElementsOfPositions(2)
    Else
        ' V0 : Dim pPositionOfTrackPoint As IJDPosition: Set pPositionOfTrackPoint = Position_FromPoint(pPointOfTrackPoint)
        Dim dDistanceMini As Double: Let dDistanceMini = BIG_EXTENSION
        Dim dDistanceMaxi As Double: Let dDistanceMaxi = -BIG_EXTENSION
        Dim i As Integer
        For i = 1 To pElementsOfPositions.Count
            Dim pPositionOfCurrentPoint As IJDPosition: Set pPositionOfCurrentPoint = pElementsOfPositions(i)
            Dim dDistance As Double: Let dDistance = pPositionOfTrackPoint.DistPt(pPositionOfCurrentPoint)
            If iTrackFlag = NEAR And dDistance < dDistanceMini Then
                Let dDistanceMini = dDistance
                Set pPositionOfResultPoint = pPositionOfCurrentPoint
            ElseIf iTrackFlag = FAR And dDistance > dDistanceMaxi Then
                Let dDistanceMaxi = dDistance
                Set pPositionOfResultPoint = pPositionOfCurrentPoint
            End If
        Next
    End If
    
    ' return result
    Set Position_FromPositionsV1 = pPositionOfResultPoint
End Function
'''Function Position_FromPositions(pElementsOfPositions As IJElements, pPointOfTrackPoint As IJPoint, ByVal iTrackFlag As Integer) As IJDPosition
'''    ' initialize result
'''    Set Position_FromPositions = Nothing
'''
'''    ' choose the nearer/further intersection point
'''    Dim pPositionOfResultPoint As IJDPosition
'''    If pPointOfTrackPoint Is Nothing Then
'''        Set pPositionOfResultPoint = pElementsOfPositions(1)
'''    Else
'''        Dim pPositionOfTrackPoint As IJDPosition: Set pPositionOfTrackPoint = Position_FromPoint(pPointOfTrackPoint)
'''        Dim dDistanceMini As Double: Let dDistanceMini = BIG_EXTENSION
'''        Dim dDistanceMaxi As Double: Let dDistanceMaxi = -BIG_EXTENSION
'''        Dim i As Integer
'''        For i = 1 To pElementsOfPositions.Count
'''            Dim pPositionOfCurrentPoint As IJDPosition: Set pPositionOfCurrentPoint = pElementsOfPositions(i)
'''            Dim dDistance As Double: Let dDistance = pPositionOfTrackPoint.DistPt(pPositionOfCurrentPoint)
'''            If iTrackFlag = NEAR And dDistance < dDistanceMini Then
'''                Let dDistanceMini = dDistance
'''                Set pPositionOfResultPoint = pPositionOfCurrentPoint
'''            ElseIf iTrackFlag = FAR And dDistance > dDistanceMaxi Then
'''                Let dDistanceMaxi = dDistance
'''                Set pPositionOfResultPoint = pPositionOfCurrentPoint
'''            End If
'''        Next
'''    End If
'''
'''    ' return result
'''    Set Position_FromPositions = pPositionOfResultPoint
'''End Function
Function Positions_FromDoubles(dDoubles() As Double, lCount As Long) As IJElements
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
Function Position_FromCoordinates(dX As Double, dY As Double, dZ As Double) As IJDPosition
    ' initialize result
    Dim pPosition As IJDPosition: Set pPosition = New DPosition
    
    Call pPosition.Set(dX, dY, dZ)
    
    ' return resultt
    Set Position_FromCoordinates = pPosition
End Function
'
' processing Vector
'
Sub Vector_Debug(pVector As IJDVector, sText As String)
    Debug.Print sText, pVector.x, pVector.y, pVector.z
End Sub
Function Vector_ToString(pVector As IJDVector, Optional sText As String = "") As String
    Let Vector_ToString = ""
    If Not sText = "" Then Let Vector_ToString = sText + ": "
    Let Vector_ToString = Vector_ToString + "u= " + CStr(pVector.x) + ", v= " + CStr(pVector.y) + ", w= " + CStr(pVector.z)
End Function
Function Vector_FromCoordinates(dU As Double, dV As Double, dW As Double) As IJDVector
    ' create new vector
    Dim pVectorNew As New DVector
    Call pVectorNew.Set(dU, dV, dW)
    
    ' return result
    Set Vector_FromCoordinates = pVectorNew
End Function
Function Vector_FromPositions(pStartPos As DPosition, pEndPos As DPosition) As IJDVector
    ' create new vector
    Dim pVectorNew As New DVector
    pVectorNew.Set pEndPos.x - pStartPos.x, pEndPos.y - pStartPos.y, pEndPos.z - pStartPos.z
    
    ' return result
    Set Vector_FromPositions = pVectorNew
End Function
Function Vector_Scale(pVector As DVector, dScale As Double) As IJDVector
    ' create new vector
    Dim pVectorNew As IJDVector
    Set pVectorNew = pVector.Clone
    
    ' set coordinates
    Call pVectorNew.Set(dScale * pVector.x, dScale * pVector.y, dScale * pVector.z)
    
    ' return result
    Set Vector_Scale = pVectorNew
End Function
Function Vector_Normalize(pVector As DVector) As IJDVector
    ' create new vector
    Dim pVectorNew As IJDVector
    Set pVectorNew = pVector.Clone
    Dim dLength As Double
    
    dLength = pVector.Length
    If dLength < EPSILON Then
        Exit Function
    End If
    ' set coordinates
    Call pVectorNew.Set(pVector.x / dLength, pVector.y / dLength, pVector.z / dLength)
    
    ' return result
    Set Vector_Normalize = pVectorNew
End Function
Function AreVectorsColinear(pVector1 As DVector, pVector2 As DVector) As Boolean
    
    If Abs(Abs(pVector1.Dot(pVector2)) - pVector1.Length() * pVector2.Length()) < EPSILON Then
        AreVectorsColinear = True
    Else
        AreVectorsColinear = False
    End If
End Function
Function AreVectorsPerpendicular(pVector1 As DVector, pVector2 As DVector) As Boolean
    
    If Abs(pVector1.Dot(pVector2)) < EPSILON Then
        AreVectorsPerpendicular = True
    Else
        AreVectorsPerpendicular = False
    End If
End Function
Function IsVectorInversionNeeded(pVector As IJDVector, pLocalCoordinateSystem As IJLocalCoordinateSystem) As Boolean
    Dim bIsVectorInversionNeeded As Boolean: Let bIsVectorInversionNeeded = False
    
    Dim pVectorOfAxis As IJDVector
    If Not pLocalCoordinateSystem Is Nothing Then
        Set pVectorOfAxis = pLocalCoordinateSystem.XAxis
    Else
        Set pVectorOfAxis = New DVector:
        Call pVectorOfAxis.Set(1, 0, 0)
    End If
    
    Dim dDot As Double:
    Let dDot = pVector.Dot(pVectorOfAxis)
    If Abs(dDot) > 0.707 Then
        If dDot < 0 Then Let bIsVectorInversionNeeded = True
    Else
        If Not pLocalCoordinateSystem Is Nothing Then
            Set pVectorOfAxis = pLocalCoordinateSystem.YAxis
        Else
            Call pVectorOfAxis.Set(0, 1, 0)
        End If
        
        Let dDot = pVector.Dot(pVectorOfAxis)
        If Abs(dDot) > 0.707 Then
            If dDot < 0 Then Let bIsVectorInversionNeeded = True
        Else
            If Not pLocalCoordinateSystem Is Nothing Then
                Set pVectorOfAxis = pLocalCoordinateSystem.ZAxis
            Else
                Call pVectorOfAxis.Set(0, 0, 1)
            End If
            
            Let dDot = pVector.Dot(pVectorOfAxis)
            If Abs(dDot) > 0.707 Then
                If dDot < 0 Then Let bIsVectorInversionNeeded = True
            End If
        End If
    End If

    Let IsVectorInversionNeeded = bIsVectorInversionNeeded
End Function
'
' processing Matrix
'
Public Function DT4x4_FromOriginAndAxes(pPositionOfOrigin As IJDPosition, _
                                        pVectorOfXAxis As IJDVector, _
                                        pVectorOfYAxis As IJDVector, _
                                        pVectorOfZAxis As IJDVector) As DT4x4
    ' fill matrix
    Dim DT4x4 As New DT4x4
    
    If True Then
        Call DT4x4.LoadIdentity

        ' set origin
        Let DT4x4.IndexValue(12) = pPositionOfOrigin.x
        Let DT4x4.IndexValue(13) = pPositionOfOrigin.y
        Let DT4x4.IndexValue(14) = pPositionOfOrigin.z
        Debug.Print "Origin", pPositionOfOrigin.x, pPositionOfOrigin.y, pPositionOfOrigin.z

        ' set x-axis
        Let DT4x4.IndexValue(0) = pVectorOfXAxis.x
        Let DT4x4.IndexValue(1) = pVectorOfXAxis.y
        Let DT4x4.IndexValue(2) = pVectorOfXAxis.z
        Debug.Print "XAxis", pVectorOfXAxis.x, pVectorOfXAxis.y, pVectorOfXAxis.z

        ' set y-axis
        Let DT4x4.IndexValue(4) = pVectorOfYAxis.x
        Let DT4x4.IndexValue(5) = pVectorOfYAxis.y
        Let DT4x4.IndexValue(6) = pVectorOfYAxis.z
        Debug.Print "YAxis", pVectorOfYAxis.x, pVectorOfYAxis.y, pVectorOfYAxis.z

        ' set zy-axis
        Let DT4x4.IndexValue(8) = pVectorOfZAxis.x
        Let DT4x4.IndexValue(9) = pVectorOfZAxis.y
        Let DT4x4.IndexValue(10) = pVectorOfZAxis.z
        Debug.Print "ZAxis", pVectorOfZAxis.x, pVectorOfZAxis.y, pVectorOfZAxis.z
    End If
    
    ' return result
    Set DT4x4_FromOriginAndAxes = DT4x4
End Function
Public Function Position_FromCoordinateSystem(pLocalCoordinateSystem As IJLocalCoordinateSystem) As IJDPosition
    Dim pPositionOfOrigin As IJDPosition
    If Not pLocalCoordinateSystem Is Nothing Then
        Set pPositionOfOrigin = pLocalCoordinateSystem.Position
    Else
        Set pPositionOfOrigin = New DPosition
        Call pPositionOfOrigin.Set(0, 0, 0)
    End If
    
    ' return result
    Set Position_FromCoordinateSystem = pPositionOfOrigin
End Function

