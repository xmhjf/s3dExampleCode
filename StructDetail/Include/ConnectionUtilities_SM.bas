Attribute VB_Name = "ConnectionUtilities_SM"
'*******************************************************************
'
'Copyright (C) 2013-2014 Intergraph Corporation. All rights reserved.
'
'File : ConnectionUtilities_SM.bas
'
'Author : Alligators
'
'Description :
'
'
'History :
'    6/Nov/2013 - vb/svsmylav
'          DI-CP-240506 'IsSeamMovement' new method is added.
'    8/Nov/2013 - vb/svsmylav
'          DI-CP-240506 'Optional gRngofChgBox' is added to 'IsSeamMovement' method.
'
'*****************************************************************************

Option Explicit
Private Const MODULE = "StructDetail\Data\Include\ConnectionUtilities_SM"

Public Const distTol = 0.0001





Public Function GetAxisCurveAtPosition(x As Double, y As Double, z As Double, oMemb As ISPSMemberPartCommon) As IJCurve
Const METHOD = "GetCurveAtEnd"
    On Error GoTo ErrorHandler
    Dim ii As Long
    Dim cnt As Long
    Dim oCurve As IJCurve
    Dim colCurves As IJElements
    Dim oCmplxCurve As IJComplexString
    
    If Not oMemb Is Nothing Then
        Set oCurve = oMemb.Axis
    Else
        Exit Function
    End If
    If Not oCurve Is Nothing Then
        If (TypeOf oCurve Is IJArc) Or (TypeOf oCurve Is IJLine) Then
            Set GetAxisCurveAtPosition = oCurve
            Exit Function
        ElseIf TypeOf oCurve Is IJComplexString Then
            Set oCmplxCurve = oCurve
            oCmplxCurve.GetCurves colCurves
            cnt = colCurves.Count
            For ii = 1 To colCurves.Count
                Set oCurve = colCurves.Item(ii)
                If oCurve.IsPointOn(x, y, z) Then
                    Set GetAxisCurveAtPosition = oCurve
                    Exit Function
                End If
            Next
        End If
    End If
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Function



Public Function GetConnectionPositionOnSupping( _
                            oSupping As ISPSMemberPartCommon, _
                            oSupped As ISPSMemberPartCommon, _
                            iEnd As Integer) As IJDPosition
Const METHOD = "GetConnectionPositionOnSupping"
    On Error GoTo ErrorHandler
    
    Dim x As Double, y As Double, z As Double
    Dim uX As Double, uY As Double, uZ As Double
    Dim oPos As IJDPosition
    Dim oLine As IJLine
    Dim transf1 As DT4x4
    Dim transf2 As DT4x4
    Dim oGeomFactory As New GeometryFactory
    
    oSupped.PointAtEnd(iEnd).GetPoint x, y, z
    oSupped.Rotation.GetTransformAtPosition x, y, z, transf1, transf2
    'ignore transf2 as this is the end of the part
    
    'get the tanget at this end of the supported member
    uX = transf1.IndexValue(0)
    uY = transf1.IndexValue(1)
    uZ = transf1.IndexValue(2)

   
    'create a line of 1 m length along the tangent
    Set oLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, x, y, z, uX, uY, uZ, 1)

   
    'we need to intersect this line with the supporting axis
    Set oPos = GetCurveLineIntersection(oSupping.Axis, oLine)
    
    
    If Not oPos Is Nothing Then
        Set GetConnectionPositionOnSupping = oPos
    End If
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function


'*************************************************************************
'Function
'
'<IsMemberAxesEndToEnd>
'
'Abstract
'
'<Determines if the Member Axis lines are connected at the ends>
'
'Arguments
'
'<Supporting and Supported Member Axis as IJLine>>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsMemberAxesEndToEnd(oLine1 As IJLine, oLine2 As IJLine) As Boolean
Const MT = "IsMemberAxesEndToEnd"
On Error GoTo ErrorHandler
    IsMemberAxesEndToEnd = False
    
    If (oLine1.length > distTol) And (oLine2.length > distTol) Then
        Dim sX1 As Double, sY1 As Double, sZ1 As Double
        Dim sX2 As Double, sY2 As Double, sZ2 As Double
        Dim eX1 As Double, eY1 As Double, eZ1 As Double
        Dim eX2 As Double, eY2 As Double, eZ2 As Double
        Dim dx As Double, dy As Double, dz As Double
        Dim dist1 As Double
    
        oLine1.GetStartPoint sX1, sY1, sZ1
        oLine1.GetEndPoint eX1, eY1, eZ1
        oLine2.GetStartPoint sX2, sY2, sZ2
        oLine2.GetEndPoint eX2, eY2, eZ2
                    
        dx = sX1 - sX2
        dy = sY1 - sY2
        dz = sZ1 - sZ2
        dist1 = Sqr(dx * dx + dy * dy + dz * dz)
        
        If Abs(dist1) < distTol Then
            GoTo Process
        Else
            dx = sX1 - eX2
            dy = sY1 - eY2
            dz = sZ1 - eZ2
            If Abs(Sqr(dx * dx + dy * dy + dz * dz)) < distTol Then
                GoTo Process
            End If
        End If
        
        dx = eX1 - sX2
        dy = eY1 - sY2
        dz = eZ1 - sZ2
        dist1 = Sqr(dx * dx + dy * dy + dz * dz)
        If Abs(dist1) < distTol Then
            GoTo Process
        Else
            dx = eX1 - eX2
            dy = eY1 - eY2
            dz = eZ1 - eZ2
            If Abs(Sqr(dx * dx + dy * dy + dz * dz)) < distTol Then
                GoTo Process
            End If
        End If
        
    End If
    
Exit Function
Process:
    IsMemberAxesEndToEnd = True
Exit Function
ErrorHandler:    HandleError MODULE, MT
End Function

 
'*************************************************************************
'Function
'
'<IsMemberAxesColinear>
'
'Abstract
'
'<Determines if the Member Axis lines are collinear>
'
'Arguments
'
'<Supporting and Supported Member Axis as IJLine>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsMemberAxesColinear(oLine1 As IJLine, oLine2 As IJLine) As Boolean
  Const MT = "IsMemberAxesColinear"
  On Error GoTo ErrorHandler
'    Currently only checking with dotproduct to see if they are in same direction

    IsMemberAxesColinear = False
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim sX1 As Double, sY1 As Double, sZ1 As Double
    Dim sX2 As Double, sY2 As Double, sZ2 As Double
    Dim eX1 As Double, eY1 As Double, eZ1 As Double
    Dim eX2 As Double, eY2 As Double, eZ2 As Double
    
    Dim oVec1 As New dVector, oVec2 As New dVector
    Dim dist1 As Double, dist2 As Double, dist3 As Double, dist4 As Double
    Dim dx As Double, dy As Double, dz As Double
    Dim cosA As Double
    Dim tol As Double

        
    oLine1.GetDirection x1, y1, z1
    oLine2.GetDirection x2, y2, z2
    
    oVec1.Set x1, y1, z1
    oVec2.Set x2, y2, z2
    cosA = oVec1.Dot(oVec2)
    tol = Abs(cosA) - 1#
 
    If Abs(tol) < distTol Then
        GoTo Process
    End If

Exit Function
Process:
    IsMemberAxesColinear = True
Exit Function
ErrorHandler:    HandleError MODULE, MT
End Function
    
'*************************************************************************
'Function
'
'<IsMemberAxesAtRightAngles>
'
'Abstract
'
'<determines if MemberAxis are at right angles>
'
'Arguments
'
'<Supporting and Supported Members as IJLine>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsMemberAxesAtRightAngles(oLine1 As IJLine, oLine2 As IJLine) As Boolean
  Const MT = "IsMemberAxesAtRightAngles"
  On Error GoTo ErrorHandler
  IsMemberAxesAtRightAngles = False

  Dim x1 As Double, y1 As Double, z1 As Double
  Dim x2 As Double, y2 As Double, z2 As Double
  Dim oVec1 As New dVector, oVec2 As New dVector
  Dim cosA As Double
  
  
  oLine1.GetDirection x1, y1, z1
  oLine2.GetDirection x2, y2, z2
  
  oVec1.Set x1, y1, z1
  'normalize vector
  oVec1.length = 1
  oVec2.Set x2, y2, z2
  'normalize vector
  oVec2.length = 1
  cosA = oVec1.Dot(oVec2)
  If Abs(cosA) < distTol Then
    IsMemberAxesAtRightAngles = True
  End If
  
  Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function


'The returned matrix would transform a point in RAD coord sys to the global coord sys
'*************************************************************************
'Function
'
'<CreateCSToMembTransform>
'
'Abstract
'
'<Creates a IJDT4x4 transformation from 2D Cross section to 3D Member coordinate system>
'
'Arguments
'
'<MemberTransformation Matrix as IJDT4x4, reflect as Boolean>
'
'Return
'
'<Transformation matrix IJDT4x4 >
'
'Exceptions
'***************************************************************************
Public Function CreateCSToMembTransform(pI4x4Mem As IJDT4x4, Optional reflect As Boolean = False) As IJDT4x4
 Const MT = "CreateCSToMembTransform"
On Error GoTo ErrorHandler

  Dim tmpI4x4 As New DT4x4
  Dim pI4x4 As New DT4x4


  tmpI4x4.LoadMatrix pI4x4Mem
  pI4x4.LoadIdentity
  'loads 0 0 -1 0 to get a local 2d CS flipped into member local CS where x along axis, Z up Y perp web
  '     -1 0  0 0
  '      0 1  0 0
  '      0 0  0 1
  'Means local x = Member -Y
  '      local y = Member Z
  '      local z = Member -X
  
  pI4x4.IndexValue(0) = 0
  If reflect = True Then
    pI4x4.IndexValue(1) = 1
  Else
    pI4x4.IndexValue(1) = -1
  End If
  pI4x4.IndexValue(5) = 0
  pI4x4.IndexValue(6) = 1
  pI4x4.IndexValue(8) = -1
  pI4x4.IndexValue(10) = 0
  
  tmpI4x4.MultMatrix pI4x4
  Set CreateCSToMembTransform = tmpI4x4
  Set pI4x4 = Nothing
Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<GetRefCollFromSmartOccurrence>
'
'Abstract
'
'<Gets the object from the ReferencesCollection Object>
'
'Arguments
'
'<SmartOccurrence as Object>
'
'Return
'
'<refcoll object As IJDReferencesCollection>
'
'Exceptions
'***************************************************************************
Public Function GetRefCollFromSmartOccurrence(pSO As IJSmartOccurrence) As IJDReferencesCollection
Const MT = "GetRefCollFromSmartOccurrence"
 On Error GoTo ErrorHandler
  
   'connect the reference collection to the smart occurrence
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = pSO
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    If Not pCollectionHelper Is Nothing Then
        If pCollectionHelper.Count = 1 Then
            Set GetRefCollFromSmartOccurrence = pCollectionHelper.Item("RC")
        End If
    End If
  
 Exit Function
ErrorHandler: HandleError MODULE, MT
End Function


'*************************************************************************
'
'<GetCurveLineIntersection>
'
'Abstract
'
'<Intersects the line with curve and returns the intersection point which is closest
'to the specified end of the line. If no end is specified, start end is assumed>
'
'Arguments
'
'<oCurve As IJCurve, oLine As IJLine, Optional bStartEnd As Boolean = True >
'
'Return
'
'<IJDPosition>
'
'Exceptions
'*************************************************************************
Public Function GetCurveLineIntersection(oCurve As IJCurve, oLine As IJLine, Optional bStartEnd As Boolean = True) As IJDPosition
Const METHOD = "GetCurveLineIntersection"
    On Error GoTo ErrorHandler
    
    Dim oStartPos As New DPosition, oEndPos As New DPosition, oInterscnPos As New DPosition
    Dim oTmpLine As IJLine
    Dim oGeomFactory As New GeometryFactory
    Dim numIntrscn As Long, numOverlaps As Long, retCode As Geom3dIntersectConstants
    Dim idx As Long
    Dim oRefPos As IJDPosition
    Dim minDistance#, dist#
    Dim x#, y#, z#, inX#, inY#, inZ#
    Dim intrscnPoints() As Double
    
    
    oLine.GetStartPoint x, y, z
    oStartPos.Set x, y, z
    
    oLine.GetEndPoint x, y, z
    oEndPos.Set x, y, z
    
    If bStartEnd = True Then
        Set oRefPos = oStartPos
    Else
        Set oRefPos = oEndPos
    End If
    
    Set oTmpLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStartPos.x, oStartPos.y, oStartPos.z, _
    oEndPos.x, oEndPos.y, oEndPos.z)
    oTmpLine.Infinite = True
    oCurve.Intersect oTmpLine, numIntrscn, intrscnPoints, numOverlaps, retCode
    If numIntrscn > 0 Then
        minDistance = 10000#  'default to a large number
        For idx = 1 To numIntrscn
            oInterscnPos.Set intrscnPoints(idx * 3 - 3), intrscnPoints(idx * 3 - 2), intrscnPoints(idx * 3 - 1)
            dist = oInterscnPos.DistPt(oRefPos)
            If dist < minDistance Then
                minDistance = dist
                Set GetCurveLineIntersection = oInterscnPos.Clone()
            End If
        Next
    Else 'no intersection
        'get minimum distance position on the curve
    
        oCurve.DistanceBetween oLine, minDistance#, x, y, z, inX, inY, inZ
        oInterscnPos.Set x, y, z
        Set GetCurveLineIntersection = oInterscnPos
    End If
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD

End Function

'***********************************************************************************************
'    Function      : IsSeamMovement
'
'    Description   : This method helps to determine if the seam is moved on given port.
'                    Currently it is intended to take bounding port of AC as input.
'
'    Parameters    :
'          Input    Bounding Port
'
'    Return        : True if the seam is moved.
'                    Optionally returns operation-id (long).
'                    and range-of-change(variant) for seam movement case (for 'optional' argument,
'                    if defined as  GBox, compilation error is noticed)
'***********************************************************************************************

Public Function IsSeamMovement(oBoundingPort As IJPort, Optional lOpn As Long = -1, Optional gRngofChgBox As Variant) As Boolean
  
  On Error GoTo ErrorHandler

    Dim sMETHOD As String
    
    sMETHOD = "IsSeamMovement"
    
    IsSeamMovement = False ': Exit Function
    
    If TypeOf oBoundingPort Is IJGeometryChangeInfo Then
        Dim oGeomChgInfo As IJGeometryChangeInfo
        Dim lOperation As Long
        Set oGeomChgInfo = oBoundingPort
        oGeomChgInfo.GetOperation lOperation
        lOpn = lOperation
        If lOperation = GEOMETRY_OPERATION_PLATE_DS_SPLIT Then
            IsSeamMovement = True
            gRngofChgBox = oGeomChgInfo.GetRange(1)     'Seam movement cases would need range of change
        End If
    End If
    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Function

