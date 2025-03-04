Attribute VB_Name = "CuttingShapes"
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : CuttingShapes.bas
'
'Author : RP
'
'Description :
'    Creates the wirebody for cutting shape and the cutback plane for  the supported member
'
'History:
'
' 09/22/03   CE             Hide any persistent construction planes
'
' 02/16/04  RP              Added a parameter, resource manager, to the
'                           cope and cutback creation methods. When resouce
'                           manager is nothing, the plane or wire body is not
'                           persisted.
'                           Moved the code that hides the plane/wirebody to the
'                           CMConstruct().
'                           Add code to take care mirror of supporting member
' 06/14/05  A. Singh        #79825. Fixed assert on trying ot locate surface from a foundation
' 08/11/05  A. Singh        TR 80543 Added call to ComputePlanarTrim to trim against a Curved Surface.
' 06/13/06  RP              Lot of changes due to impact from curved members. Modified some existing methods
'                           to take care of impacts (DI#84001)
' 11/02/06  RP              TR#108733 and DI#84749 - Copelength and depth not computed. Also
'                           moved compute of these properties of aggregator section to avoid
'                           extra compute
'
'8/10/2007  RP              TR#124630 - Resolved error while placing assembly connection between members
'                           with tubular section.Added code in GetCopeDimension()  to check section type
'                           of supporting member. If the section type is circular then the method exits without
'                           computing cope dimensions.Anyway cope dimension doesn't make sense when we are
'                           cutting against a member with circular section
'26/05/16  APK              TR-294293: Validation Checks were added to avoid record exceptions in GetCopeDimensions()
'*******************************************************************************************************

Option Explicit
Private Const MODULE = "CuttingShapes"
Private Const PI As Double = 3.1415926

Public Sub CreateTopCopeforShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, _
        iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)

 Const MT = "CreateTopCopeforShape"
 On Error GoTo ErrorHandler
 
 Dim oProfile As IJCrossSection
 Dim SquareEnd As Boolean
 
  If Not oSupping Is Nothing Then
        Dim OSPSCrossSection As ISPSCrossSection
        Set OSPSCrossSection = oSupping.CrossSection
        If Not OSPSCrossSection Is Nothing Then
            Set oProfile = OSPSCrossSection.definition
            If Not oProfile Is Nothing Then
                Dim strSectionType As String
                strSectionType = oProfile.Type
            End If
        End If
    End If
 SquareEnd = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("SquaredEnd").Value
 Select Case strSectionType
    Case "HSSC", "CS", "PIPE"
        CreateTopCopeforRoundShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oCutter
    Case "W", "S", "HP", "M"
        If Not SquareEnd Then
            CreateTopCopeforWShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oCutter
        Else
            CreateSquaredCopeForWShape oSupping, oSupped, iEnd, oAttribsCAO, True, oRescMgr, oCutter
        End If
    Case "L"
         'No top cope for L
    Case "C", "MC"
       CreateTopCopeforCShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oCutter
    Case "WT", "MT", "ST"
        CreateTopCopeforTShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oCutter
    Case "2L"
        'no top cope for 2L
    Case "RS", "HSSR"
        'no cope for rectangular shapes
    Case Else
      'Err.Raise
 End Select
    
 Set oProfile = Nothing
  
 Exit Sub
ErrorHandler:         HandleError MODULE, MT
  End Sub

Public Sub CreateBottomCopeforShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, _
        iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)

 Const MT = "CreateBottomCopeforShape"
 On Error GoTo ErrorHandler
 
 Dim oProfile As IJCrossSection
 Dim SquareEnd As Boolean
  
 If Not oSupping Is Nothing Then
        Dim OSPSCrossSection As ISPSCrossSection
        Set OSPSCrossSection = oSupping.CrossSection
        If Not OSPSCrossSection Is Nothing Then
            Set oProfile = OSPSCrossSection.definition
            If Not oProfile Is Nothing Then
                Dim strSectionType As String
                strSectionType = oProfile.Type
            End If
        End If
    End If
 SquareEnd = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("SquaredEnd").Value
 Select Case strSectionType
    Case "HSSC", "CS", "PIPE"
        'no bottom cope for round shapes
    Case "W", "S", "HP", "M"
        If Not SquareEnd Then
            CreateBottomCopeforWShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oCutter
        Else
            CreateSquaredCopeForWShape oSupping, oSupped, iEnd, oAttribsCAO, False, oRescMgr, oCutter
        End If
    Case "L"
        CreateBottomCopeforLShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oCutter
    Case "C", "MC"
        CreateBottomCopeforCShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oCutter
    Case "WT", "MT", "ST"
        'No bottom cope needed for T
    Case "2L"
        CreateBottomCopefor2LShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oCutter
    Case "RS", "HSSR"
        'no cope for rectangular shapes
    Case Else
      'Err.Raise
 End Select
    
 Set oProfile = Nothing
  
 Exit Sub
ErrorHandler:         HandleError MODULE, MT
  End Sub
Public Sub GetCutbackSurface(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, _
        iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oSurface As IJSurface)
    Const MT = "GetCutbackSurface"
    On Error GoTo ErrorHandler
    
    Dim SquareEnd As Boolean
    Dim oProfile As IJCrossSection
    
     If Not oSupping Is Nothing Then
        Dim OSPSCrossSection As ISPSCrossSection
        Set OSPSCrossSection = oSupping.CrossSection
        If Not OSPSCrossSection Is Nothing Then
            Set oProfile = OSPSCrossSection.definition
            If Not oProfile Is Nothing Then
                Dim strSectionType As String
                strSectionType = oProfile.Type
            End If
        End If
    End If
    'If cutback is not created as part cope then we need to handle the folowing not to throw an error
    SquareEnd = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("SquaredEnd").Value
    Select Case strSectionType
        Case "W", "S", "HP", "M"
            If Not SquareEnd Then
                CreateCutback oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oSurface
            Else
                CreateSquaredCutbackForWShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oSurface
            End If
        Case Else
            CreateCutback oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oSurface
    End Select

    Exit Sub
ErrorHandler:         HandleError MODULE, MT
  End Sub


Public Sub CreateCutbackForShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, _
        iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oSurface As IJSurface)
    Const MT = "CreateCutbackForShape"
    On Error GoTo ErrorHandler
    
    Dim SquareEnd As Boolean
    Dim bAlwaysPlanar As Boolean
    Dim iQuadrant As Integer
    Dim oProfile As IJCrossSection
    Dim bIsCopeNeeded As Boolean, bIsCope1Needed As Boolean, bIsCope2Needed As Boolean
    Dim oAxisCurve As IJCurve
    Dim oPlane As IJPlane
    
    iQuadrant = GetIncidentMemberQuadrant(oSupping, oSupped, iEnd)
    bIsCope1Needed = IsCope1NeededByShapeAndIncidence(oSupping, iQuadrant)
    bIsCope2Needed = IsCope2NeededByShapeAndIncidence(oSupping, iQuadrant)
    
    If bIsCope1Needed Or bIsCope2Needed Then
    bIsCopeNeeded = True
    End If
        
    If Not oSupping Is Nothing Then
        Dim OSPSCrossSection As ISPSCrossSection
        Set OSPSCrossSection = oSupping.CrossSection
        If Not OSPSCrossSection Is Nothing Then
            Set oProfile = OSPSCrossSection.definition
            If Not oProfile Is Nothing Then
                Dim strSectionType As String
                strSectionType = oProfile.Type
            End If
        End If
    End If
    'If cutback is not created as part cope then we need to handle the folowing not to throw an error
    On Error Resume Next
        SquareEnd = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("SquaredEnd").Value
        bAlwaysPlanar = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("AlwaysPlanar").Value
    On Error GoTo ErrorHandler
    If bAlwaysPlanar Or (Not bIsCopeNeeded) Then
        If Not SquareEnd Then
            If Not oSurface Is Nothing Then
                If Not TypeOf oSurface Is IJSurfaceBody Then
                    CreateCutbackForRectShape oSupping, oSupped, iEnd, oAttribsCAO, Nothing, oSurface
                    Exit Sub
                End If
            End If
            CreateCutbackForRectShape oSupping, oSupped, iEnd, oAttribsCAO, Nothing, oPlane
            
            'from the supping axis and plane, create a surface of projection or revolution
            CreateCutbackSurface oRescMgr, oSupping, oSupped, iEnd, oPlane, oSurface

        Else
            CreateSquaredCutbackForRectShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oSurface
        End If
    Else
        Select Case strSectionType
            Case "W", "S", "HP", "M"
                If Not SquareEnd Then
                    If Not oSurface Is Nothing Then
                        If Not TypeOf oSurface Is IJSurfaceBody Then
                            CreateCutback oSupping, oSupped, iEnd, oAttribsCAO, Nothing, oSurface
                            Exit Sub
                        End If
                    End If
                    
                    CreateCutback oSupping, oSupped, iEnd, oAttribsCAO, Nothing, oPlane
                    'from the supping axis and plane, create a surface of projection or revolution
                    CreateCutbackSurface oRescMgr, oSupping, oSupped, iEnd, oPlane, oSurface
                
                Else
                    CreateSquaredCutbackForWShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oSurface
                End If
            Case Else
                If Not oSurface Is Nothing Then
                    If Not TypeOf oSurface Is IJSurfaceBody Then
                        CreateCutback oSupping, oSupped, iEnd, oAttribsCAO, Nothing, oSurface
                        Exit Sub
                    End If
                End If
                
                CreateCutback oSupping, oSupped, iEnd, oAttribsCAO, Nothing, oPlane
                'from the supping axis and plane, create a surface of projection or revolution
                CreateCutbackSurface oRescMgr, oSupping, oSupped, iEnd, oPlane, oSurface
        
        End Select
    End If
    Set oProfile = Nothing
    Exit Sub
ErrorHandler:         HandleError MODULE, MT
  End Sub
Public Sub GetCopeDimensions(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, ByRef dblLength As Double, ByRef dblDepth As Double)
Const MT = "GetCopeDimensions"
    Dim oSurface As IJSurface
    Dim oCornerPos As IJDPosition
    Dim oElms As IJElements
    Dim code As Geom3dIntersectConstants
    Dim dblIncrement As Double
    Dim oCopePlane As IJPlane
    Dim oNearPoint As IJPoint, oFarPoint As IJPoint
    Dim oCopePoint As IJPoint
    Dim x#, y#, z#, nx#, ny#, nz#, x1#, y1#, z1#
    Dim pGeometryFactory As New GeometryFactory
    Dim oVec As New DVector
    Dim oLengthLine As IJLine
    Dim strSecType As String
    Dim oAttribsAC As IJDAttributesCol
    Dim oDepthVec As IJDVector, oLengthVec As IJDVector
    Dim bSquareEnd As Boolean
    Dim oPosAlong As IJDPosition
    Dim bTopCope As Boolean, bIsCope1Needed As Boolean, bIsCope2Needed As Boolean, bCircularCope As Boolean
    Dim iQuadrant As Integer
    Dim oProfileBO As ISPSCrossSection
    Dim strProfileType As String
    

    iQuadrant = GetIncidentMemberQuadrant(oSupping, oSupped, iEnd)
    bIsCope1Needed = IsCope1NeededByShapeAndIncidence(oSupping, iQuadrant)
    bIsCope2Needed = IsCope2NeededByShapeAndIncidence(oSupping, iQuadrant)
    
    If Not oSupping Is Nothing Then
        Set oProfileBO = oSupping.CrossSection
        If Not oProfileBO Is Nothing Then
            Dim oCrossection As IJCrossSection
            Set oCrossection = oProfileBO.definition
            If Not oCrossection Is Nothing Then
                strProfileType = oCrossection.Type
            End If
        End If
    End If
    bCircularCope = False
    If (strProfileType = "HSSC") Or (strProfileType = "CS") Or (strProfileType = "PIPE") Then
        bCircularCope = True
    End If
    
    If (Not bIsCope1Needed And Not bIsCope2Needed) Or bCircularCope Then
        dblLength = 0
        dblDepth = 0
        Exit Sub
    End If

    GetCutbackSurface oSupping, oSupped, iEnd, oAttribsCAO, Nothing, oSurface
    GetCopeCornerPosition oSupping, oSupped, iEnd, oAttribsCAO, Nothing, oCornerPos


    '
    '  TopCope
    '
    '        |
    '        |Depth
    '        |
    '  ------ oCornerPos
    '  Length
    
    '
    '  BottomCope
    '
    '  Length
    '  -------- oCornerPos
    '          |
    '          | Depth
    '          |
    '
    Set oAttribsAC = oAttribsCAO.CollectionOfAttributes("IJUASPSCope")
    bSquareEnd = oAttribsAC.Item("SquaredEnd").Value
    
    'create line passiong through oCornerPos
    Set oPosAlong = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)
    If bSquareEnd Then
        Set oLengthVec = GetXVector(oSupped, oPosAlong)
        Set oDepthVec = GetXVector(oSupping, oPosAlong)
        Set oDepthVec = oDepthVec.Cross(oLengthVec)
        oDepthVec.length = 1
        bTopCope = IsSupportedAxisInPositiveZ(oSupping, oSupped, iEnd)
    Else
        Set oLengthVec = GetYVector(oSupping, oPosAlong)
        Set oDepthVec = GetZVector(oSupping, oPosAlong)
        oDepthVec.length = 1
        bTopCope = bIsCope1Needed
    End If
    oCornerPos.Get x, y, z
    oLengthVec.Get nx, ny, nz
    Set oCopePlane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, x, y, z, nx, ny, nz)
    
    Set oLengthLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, x, y, z, nx, ny, nz, 10#)
   
    oLengthLine.Infinite = True
    oSurface.Intersect oLengthLine, oElms, code
    
    dblIncrement = oAttribsAC.Item("Increment").Value
    If Not oElms Is Nothing Then
        If oElms.count > 0 Then
            Set oFarPoint = oElms.Item(1) ' the intersection point
            Set oNearPoint = pGeometryFactory.Points3d.CreateByPoint(Nothing, x, y, z)
            'based on the section range box of the supported, get the top most or bottom most intersection point
            'on the cope plane. bTopCope=True for Top, Bottom otherwise
            GetCopePoint oSupping, oSupped, iEnd, oCopePlane, bTopCope, oCopePoint
            'calculate cope length
            dblLength = oFarPoint.DistFromPt(oNearPoint)
            oCopePoint.GetPoint x1, y1, z1
            oVec.Set x1 - x, y1 - y, z1 - z
            
            'calculate cope depth
            dblDepth = Abs(oDepthVec.Dot(oVec))
            dblLength = RoundOff(dblIncrement, dblLength)
            dblDepth = RoundOff(dblIncrement, dblDepth)
        End If
    End If
    Exit Sub
End Sub

Public Sub GetCopeCornerPosition(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCornerPos As IJDPosition)
Const MT = "GetCopeCornerPosition"

    Dim depth As Double, tFlange As Double, tWeb As Double, bFlange As Double
    Dim cp As Long
    Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double
    Dim xOffsetCP5 As Double, yOffsetCP5 As Double
    Dim oMat As New DT4x4
    Dim oVec As New DVector
    Dim oSectionAttrbs As IJDAttributes
    Dim oProfileBO As ISPSCrossSection
    Dim FlangeClearance1#, FlangeClearance2#
    Dim oPosAlongAxis As IJDPosition
    Dim iQuadrant As Long
    Dim SquareEnd As Boolean
    Dim oAttribsAC As IJDAttributesCol, oAttribsSection As IJDAttributesCol
    Dim oPosAlong As IJDPosition
    Dim bMirror As Boolean
    
    
    Set oAttribsAC = oAttribsCAO.CollectionOfAttributes("IJUASPSCope")
    
    SquareEnd = oAttribsAC.Item("SquaredEnd").Value
    
    FlangeClearance1 = oAttribsAC.Item("FlangeClearance1").Value
    FlangeClearance2 = oAttribsAC.Item("FlangeClearance2").Value
    
    If Not oSupping Is Nothing Then
        Set oProfileBO = oSupping.CrossSection
        If Not oProfileBO Is Nothing Then
        Dim oCrossection As IJCrossSection
        Set oCrossection = oProfileBO.definition
            If Not oCrossection Is Nothing Then
                Dim strSectionType As String
                strSectionType = oCrossection.Type
            End If
        End If
    End If
    Set oSectionAttrbs = oSupping.CrossSection.definition
    
    Set oAttribsSection = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions")
    depth = oAttribsSection.Item("Depth").Value
    bFlange = oAttribsSection.Item("Width").Value
    tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
    
    iQuadrant = GetIncidentMemberQuadrant(oSupping, oSupped, iEnd)
    bMirror = oSupping.Rotation.Mirror
    
    Set oCornerPos = New DPosition
    Select Case strSectionType
        Case "W", "S", "HP", "M"
            If SquareEnd Then
                GetSquaredCopeCornerPosition oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oCornerPos
                'the position is already in global coordinates, so exit
                Exit Sub
            Else
                oCornerPos.x = (bFlange / 2# + FlangeClearance1)
                oCornerPos.y = (depth / 2#) - (tFlange + FlangeClearance2)
                oCornerPos.z = 0
                If (iQuadrant = 4 And Not bMirror) Or (iQuadrant = 2 And bMirror) Then
                    oCornerPos.x = -1 * oCornerPos.x
                End If
            End If
        Case "L"
            If (iQuadrant = 4 And bMirror) Or (iQuadrant = 2 And Not bMirror) Then
                oCornerPos.x = (bFlange / 2# + FlangeClearance1)
                oCornerPos.y = -1 * ((depth / 2#) - (tFlange + FlangeClearance2))
                oCornerPos.z = 0
            ElseIf iQuadrant = 1 Then
                oCornerPos.x = -(bFlange / 2# - tFlange - FlangeClearance2)
                oCornerPos.y = -((depth / 2#) + FlangeClearance1)
                oCornerPos.z = 0
            End If
        Case "C", "MC"
            oCornerPos.x = (bFlange / 2# + FlangeClearance1)
            oCornerPos.y = (depth / 2#) - (tFlange + FlangeClearance2)
            oCornerPos.z = 0
        Case "WT", "MT", "ST"
            oCornerPos.x = (bFlange / 2# + FlangeClearance1)
            oCornerPos.y = (depth / 2#) - (tFlange + FlangeClearance2)
            oCornerPos.z = 0
            If (iQuadrant = 4 And Not bMirror) Or (iQuadrant = 2 And bMirror) Then
                oCornerPos.x = -1 * oCornerPos.x
            End If
        Case "2L"
            oCornerPos.x = (bFlange / 2# + FlangeClearance1)
            oCornerPos.y = -1 * ((depth / 2#) - (tFlange + FlangeClearance2))
            oCornerPos.z = 0
            If (iQuadrant = 4 And Not bMirror) Or (iQuadrant = 2 And bMirror) Then
                oCornerPos.x = -1 * oCornerPos.x
            End If
        Case Else
            'for any unknown cross section, retun ERROR
            Err.Raise E_FAIL
    End Select
    cp = oProfileBO.CardinalPoint
    oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
    oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.
    
    xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
    yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
    
    oVec.Set xOffset, yOffset, 0
    oMat.LoadIdentity
    oMat.Translate oVec
    Set oCornerPos = oMat.TransformPosition(oCornerPos)
    Set oPosAlong = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)
    'transform to global
    If Not oPosAlong Is Nothing Then
        oSupping.Rotation.GetTransformAtPosition oPosAlong.x, oPosAlong.y, oPosAlong.z, oMat, Nothing
    Else
        oSupping.Rotation.GetTransform oMat
    End If
    Set oMat = CreateCSToMembTransform(oMat, oSupping.Rotation.Mirror)
    Set oCornerPos = oMat.TransformPosition(oCornerPos)

    Exit Sub
End Sub
Public Sub GetSquaredCopeCornerPosition(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCornerPos As IJDPosition)
Const MT = "GetSquaredCopeCornerPosition"
    Dim oCornerLines() As IJLine
    Dim oInterSecPoints() As IJDPosition

    Dim oPos As New DPosition
    Dim x#, y#, z#, nx#, ny#, nz#
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim oMat As IJDT4x4
    Dim oSectionAttrbs As IJDAttributes
    Dim oLeftBottomOfTopFlange As New DPosition, oRightBottomOfTopFlange As New DPosition
    Dim oLeftTopOfBottomFlange As New DPosition, oRightTopOfBottomFlange As New DPosition
    Dim FlangeClearance1#, FlangeClearance2#
    Dim depth As Double, tFlange As Double, bFlange As Double
    Dim oFlangeSidePlane As IJPlane
    Dim pGeometryFactory As New GeometryFactory
    Dim oVec As IJDVector
    Dim bIsSuppedAxisInPositiveY As Boolean, bIsSuppedAxisInPositiveZ As Boolean
    Dim colElms As IJElements
    Dim oIntrSecPoint As IJPoint
    Dim oCopePlane As IJPlane
    Dim oSuppingZvec As IJDVector
    Dim oCopeYvec As IJDVector, oCopeXvec As New DVector, oCopeZvec As IJDVector
    Dim oIntersecPos As New DPosition
    Dim oPosAlongAxis As IJDPosition
    Dim oSurface As IJSurface
    Dim oAttribsAC As IJDAttributesCol, oAttribsSection As IJDAttributesCol
    
    
    Set oAttribsAC = oAttribsCAO.CollectionOfAttributes("IJUASPSCope")
    FlangeClearance1 = oAttribsAC.Item("FlangeClearance1").Value
    FlangeClearance2 = oAttribsAC.Item("FlangeClearance2").Value
    
    
    Set oSectionAttrbs = oSupping.CrossSection.definition
    Set oAttribsSection = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions")
  
    depth = oAttribsSection.Item("Depth").Value
    bFlange = oAttribsSection.Item("Width").Value
    tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
    
    Set oPosAlongAxis = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)
    'Set flange points
    oLeftBottomOfTopFlange.Set -(bFlange / 2 + FlangeClearance1), (depth / 2 - tFlange - FlangeClearance2), 0
    oRightBottomOfTopFlange.Set (bFlange / 2 + FlangeClearance1), (depth / 2 - tFlange - FlangeClearance2), 0
    oLeftTopOfBottomFlange.Set -(bFlange / 2 + FlangeClearance1), -(depth / 2 - tFlange - FlangeClearance2), 0
    oRightTopOfBottomFlange.Set (bFlange / 2 + FlangeClearance1), -(depth / 2 - tFlange - FlangeClearance2), 0
    

    'Flange points are  global now
    Set oLeftBottomOfTopFlange = TransformPosToGlobal(oSupping, oLeftBottomOfTopFlange, oPosAlongAxis)
    Set oRightBottomOfTopFlange = TransformPosToGlobal(oSupping, oRightBottomOfTopFlange, oPosAlongAxis)
    Set oLeftTopOfBottomFlange = TransformPosToGlobal(oSupping, oLeftTopOfBottomFlange, oPosAlongAxis)
    Set oRightTopOfBottomFlange = TransformPosToGlobal(oSupping, oRightTopOfBottomFlange, oPosAlongAxis)
    
    'now create a plane with either left or right flnage point as the root point and supping Y axis as normal
    'This plane represents the bounding plane  either the left or right side.
    If IsSupportedAxisInPositiveY(oSupping, oSupped, iEnd) = True Then
        bIsSuppedAxisInPositiveY = True
    End If
    If IsSupportedAxisInPositiveZ(oSupping, oSupped, iEnd) = True Then
        bIsSuppedAxisInPositiveZ = True
    End If
    If bIsSuppedAxisInPositiveY Then
        If oSupping.Rotation.Mirror = False Then
            oLeftBottomOfTopFlange.Get x, y, z
        Else
            oRightBottomOfTopFlange.Get x, y, z
        End If
    Else
        If oSupping.Rotation.Mirror = False Then
            oRightBottomOfTopFlange.Get x, y, z
        Else
            oLeftBottomOfTopFlange.Get x, y, z
        End If
    End If
    
    Set oVec = GetYVector(oSupping, oPosAlongAxis)
    oVec.Get nx, ny, nz
    Set oFlangeSidePlane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, x, y, z, nx, ny, nz)
    
    'from the supping axis and plane, create a surface of projection or revolution
    CreateCutbackSurface Nothing, oSupping, oSupped, iEnd, oFlangeSidePlane, oSurface
    
    'get lines connecting cross section corner points
    GetCrossSectionLines oSupped, iEnd, oCornerLines
    
    'now intersect the cross section lines with the flange side  plane
    GetIntersectionPoints oCornerLines, oSurface, oInterSecPoints
    
    'get intersection point that is farthest from the end being connected.
    'that is closest to other end
    Set oIntersecPos = GetPosClosestToOtherEnd(oInterSecPoints, oCornerLines, iEnd)

    oIntersecPos.Get x, y, z
    'get supped axis direction
    oSupped.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    If iEnd = SPSMemberAxisStart Then
        oSupped.Rotation.GetTransformAtPosition sX, sY, sZ, oMat, Nothing
        oCopeXvec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    Else
        oSupped.Rotation.GetTransformAtPosition eX, eY, eZ, oMat, Nothing
        oCopeXvec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
        oCopeXvec.[Scale] -1 'reverse direction such that the vector is away from the connected end
    End If

    oCopeXvec.length = 1
    
    Set oSuppingZvec = GetZVector(oSupping, oPosAlongAxis)
    Set oCopeZvec = oCopeXvec.Cross(oSuppingZvec)
    oCopeZvec.length = 1
    Set oCopeYvec = oCopeZvec.Cross(oCopeXvec)
    oCopeYvec.length = 1
    
    oCopeXvec.Get nx, ny, nz
    
    'create plane with the intersection point(that gives max cope distance) as root point and supped x axis as the normal
    'this plane is parallel to the cutback plane but offset by the cope length
    Set oCopePlane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, x, y, z, nx, ny, nz)
    
    oSupping.Rotation.GetTransformAtPosition oPosAlongAxis.x, oPosAlongAxis.y, oPosAlongAxis.z, oMat, Nothing
    oMat.Invert 'global to local
    'transform intersection point to supping local coordinates
    Set oPos = oMat.TransformPosition(oIntersecPos)
    oPos.Get x, y, z
    
    
    'transform flange point to supping local coordinates
    Set oLeftBottomOfTopFlange = oMat.TransformPosition(oLeftBottomOfTopFlange)
    Set oLeftTopOfBottomFlange = oMat.TransformPosition(oLeftTopOfBottomFlange)
    Set oRightBottomOfTopFlange = oMat.TransformPosition(oRightBottomOfTopFlange)
    Set oRightTopOfBottomFlange = oMat.TransformPosition(oRightTopOfBottomFlange)
    
    'change position of flange points such that they are at the same local x as the intersection point
    oLeftBottomOfTopFlange.x = x
    oLeftTopOfBottomFlange.x = x
    oRightBottomOfTopFlange.x = x
    oRightTopOfBottomFlange.x = x
    oMat.Invert 'local to global
    
    Set oLeftBottomOfTopFlange = oMat.TransformPosition(oLeftBottomOfTopFlange)
    Set oLeftTopOfBottomFlange = oMat.TransformPosition(oLeftTopOfBottomFlange)
    Set oRightBottomOfTopFlange = oMat.TransformPosition(oRightBottomOfTopFlange)
    Set oRightTopOfBottomFlange = oMat.TransformPosition(oRightTopOfBottomFlange)
    
    If bIsSuppedAxisInPositiveY Then
        If bIsSuppedAxisInPositiveZ Then
            Set oCornerPos = oLeftBottomOfTopFlange
        Else
            Set oCornerPos = oLeftTopOfBottomFlange
        End If
    Else
        If bIsSuppedAxisInPositiveZ Then
            Set oCornerPos = oRightBottomOfTopFlange
        Else
            Set oCornerPos = oRightTopOfBottomFlange
        End If
    End If
    Exit Sub
End Sub


'Used for any shape. The CMConstruct() for cope uses this. The CMEvaluate() will call shape specific cope
'creation method. The CMConstruct calling this lighter version makes it faster. Also reduces error
'as the specific cope creation method may require the cutback and it may not be computed when the cope
'is constructed
Public Sub CreateSimpleTopCopeforShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
Const MT = "CreateSimpleTopCopeforShape"
  On Error GoTo ErrorHandler

  Dim depth As Double, bFlange As Double
  Dim cp As Long
  Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
  Dim oLine3d(1 To 4) As Line3d
  Dim pGeometryFactory As New GeometryFactory
  Dim oPos(1 To 4) As IJDPosition
  Dim idx As Integer
  Dim oCmplx As ComplexString3d
  Dim oMat As New DT4x4
  Dim oVec As New DVector
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim oArrayOfids(12) As Long
  Dim oSectionAttrbs As IJDAttributes
  Dim oProfileBO As ISPSCrossSection

     
  Set oProfileBO = oSupping.CrossSection
  Set oSectionAttrbs = oSupping.CrossSection.definition
  depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
  bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value

  
  cp = oProfileBO.CardinalPoint
  oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
  oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.

  xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
  yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
  
  For idx = 1 To 4
    Set oPos(idx) = New DPosition
  Next idx
  
  'CP7 - oPos(1) CP9 - oPos(2) going clockwise
 
  oPos(1).x = -1 * (bFlange / 2#)
  oPos(1).y = (depth / 2#)
  oPos(1).z = 0

  oPos(2).x = -1 * oPos(1).x
  oPos(2).y = oPos(1).y
  oPos(2).z = 0

  oPos(3).x = oPos(2).x
  oPos(3).y = 0
  oPos(3).z = 0

  oPos(4).x = -1 * oPos(3).x
  oPos(4).y = oPos(3).y
  oPos(4).z = 0
  ' Make edges away from the supported member longer so that we don't get any hanging pieces
  oPos(1).y = oPos(1).y + LARGE_EDGE
  oPos(2).y = oPos(2).y + LARGE_EDGE


  
  Dim curveElms As IJElements
  Set curveElms = New JObjectCollection
  
  For idx = 1 To 4
    Dim nextIdx As Integer
    nextIdx = idx Mod 4
    Set oLine3d(idx) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, _
    oPos(nextIdx + 1).x, oPos(nextIdx + 1).y, oPos(nextIdx + 1).z)
    
    curveElms.Add oLine3d(idx)
    Set oLine3d(idx) = Nothing
  Next idx

  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
  Set curveElms = Nothing


  oVec.Set xOffset, yOffset, 0
  
  oMat.LoadIdentity
  oMat.Translate oVec
  oCmplx.Transform oMat ' This positions the cutting shape with the current CP as origin. We assumed CP5 as origin while
  'creating the shape

  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter

  For idx = 1 To 4
    Set oPos(idx) = Nothing
  Next idx
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub
'Used for any shape. The CMConstruct() for cope uses this. The CMEvaluate() will call shape specific cope
'creation method. The CMConstruct calling this lighter version makes it faster. Also reduces error
'as the specific cope creation method may require the cutback and it may not be computed when the cope
'is constructed
Public Sub CreateSimpleBottomCopeforShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
Const MT = "CreateSimpleBottomCopeforShape"
  On Error GoTo ErrorHandler

  Dim depth As Double, bFlange As Double
  Dim cp As Long
  Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
  Dim oLine3d(1 To 4) As Line3d
  Dim pGeometryFactory As New GeometryFactory
  Dim oPos(1 To 4) As IJDPosition
  Dim idx As Integer
  Dim oCmplx As ComplexString3d
  Dim oMat As New DT4x4
  Dim oVec As New DVector
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim oArrayOfids(12) As Long
  Dim oSectionAttrbs As IJDAttributes
  Dim oProfileBO As ISPSCrossSection

     
  Set oProfileBO = oSupping.CrossSection
  Set oSectionAttrbs = oSupping.CrossSection.definition
  depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
  bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value

  
  cp = oProfileBO.CardinalPoint
  oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
  oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.

  xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
  yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
  
  For idx = 1 To 4
    Set oPos(idx) = New DPosition
  Next idx
  
  'CP1 - oPos(1) CP3 - oPos(2) going anti clockwise
 
  oPos(1).x = -1 * (bFlange / 2#)
  oPos(1).y = -1 * (depth / 2#)
  oPos(1).z = 0

  oPos(2).x = -1 * oPos(1).x
  oPos(2).y = oPos(1).y
  oPos(2).z = 0

  oPos(3).x = oPos(2).x
  oPos(3).y = 0
  oPos(3).z = 0

  oPos(4).x = -1 * oPos(3).x
  oPos(4).y = oPos(3).y
  oPos(4).z = 0
  ' Make edges away from the supported member longer so that we don't get any hanging pieces
  oPos(1).y = oPos(1).y - LARGE_EDGE
  oPos(2).y = oPos(2).y - LARGE_EDGE

  
  Dim curveElms As IJElements
  Set curveElms = New JObjectCollection
  
  For idx = 1 To 4
    Dim nextIdx As Integer
    nextIdx = idx Mod 4
    Set oLine3d(idx) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, _
    oPos(nextIdx + 1).x, oPos(nextIdx + 1).y, oPos(nextIdx + 1).z)
    
    curveElms.Add oLine3d(idx)
    Set oLine3d(idx) = Nothing
  Next idx

  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
  Set curveElms = Nothing


  oVec.Set xOffset, yOffset, 0
  
  oMat.LoadIdentity
  oMat.Translate oVec
  oCmplx.Transform oMat ' This positions the cutting shape with the current CP as origin. We assumed CP5 as origin while
  'creating the shape


  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter
  
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub

Private Sub CreateTopCopeforRoundShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, _
        iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
  Const MT = "CreateTopCopeforRoundShape"
  On Error GoTo ErrorHandler
  
  
  Dim oSectionAttrbs As IJDAttributes
  Dim oProfileBO As ISPSCrossSection
  Dim depth As Double, dblRadius As Double
  Dim cp As Long
  Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
  Dim oOrigin As New DPosition
  Dim oNormal As New DVector, oVec As New DVector
  Dim pGeometryFactory As New GeometryFactory
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oCircle3d As Circle3d
  Dim oMat As New DT4x4
  Dim oArrayOfids(1) As Long
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim FlangeClearance#
  Dim oCmplx As ComplexString3d
  Dim colCurves As IJElements

  
  
  FlangeClearance = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance1").Value
  
  Set oProfileBO = oSupping.CrossSection
  Set oSectionAttrbs = oSupping.CrossSection.definition
  
  depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
  dblRadius = depth / 2 + FlangeClearance
  
  cp = oProfileBO.CardinalPoint

  oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
  oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.

  xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
  yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
  
  oOrigin.Set 0, 0, 0
  oNormal.Set 0, 0, 1
  Set oCircle3d = pGeometryFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                    oOrigin.x, oOrigin.y, oOrigin.z, oNormal.x, oNormal.y, oNormal.z, dblRadius)
  oVec.Set xOffset, yOffset, 0
  
  oMat.LoadIdentity
  oMat.Translate oVec
  oCircle3d.Transform oMat ' This positions the cutting shape with the current CP as origin. We assumed CP5 as origin while
  'creating the shape
  Set colCurves = New JObjectCollection
  colCurves.Add oCircle3d
  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, colCurves)
 
  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter

  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub

Public Sub CreateTopCopeforWShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
Const MT = "CreateTopCopeforWShape"
  On Error GoTo ErrorHandler


  Dim depth As Double, tFlange As Double, tWeb As Double, bFlange As Double
  Dim cp As Long
  Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
  Dim oLine3d(1 To 4) As Line3d
  Dim pGeometryFactory As New GeometryFactory
  Dim oPos(1 To 4) As IJDPosition
  Dim idx As Long
  Dim oCmplx As ComplexString3d
  Dim oMat As New DT4x4
  Dim oVec As New DVector
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim oArrayOfids(12) As Long
  Dim oSectionAttrbs As IJDAttributes
  Dim oProfileBO As ISPSCrossSection
  Dim WebClearance#, FlangeClearance1#, FlangeClearance2#

  WebClearance = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("WebClearance").Value
  FlangeClearance1 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance1").Value
  FlangeClearance2 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance2").Value
    
  Set oProfileBO = oSupping.CrossSection
  Set oSectionAttrbs = oSupping.CrossSection.definition
  
  depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
  bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
  tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
  tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
  
  cp = oProfileBO.CardinalPoint

  oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
  oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.

  xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
  yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
  
  For idx = 1 To 4
    Set oPos(idx) = New DPosition
  Next idx
  
  'CP7 - oPos(1) CP9 - oPos(2)
  

  oPos(1).x = -1 * (bFlange / 2# + FlangeClearance1)
  oPos(1).y = (depth / 2#) + FlangeClearance1
  oPos(1).z = 0

  oPos(2).x = -1 * oPos(1).x
  oPos(2).y = oPos(1).y
  oPos(2).z = 0

  oPos(3).x = oPos(2).x
  oPos(3).y = (depth / 2#) - (tFlange + FlangeClearance2)
  oPos(3).z = 0

  oPos(4).x = -1 * oPos(3).x
  oPos(4).y = oPos(3).y
  oPos(4).z = 0
  
  
  ' Make edges away from the supported member longer so that we don't get any hanging pieces
  oPos(1).y = oPos(1).y + LARGE_EDGE
  oPos(2).y = oPos(2).y + LARGE_EDGE


  
  Dim curveElms As IJElements
  Set curveElms = New JObjectCollection
  
  
  For idx = 1 To 4
    Dim nextIdx As Integer
    nextIdx = idx Mod 4
    Set oLine3d(idx) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, _
    oPos(nextIdx + 1).x, oPos(nextIdx + 1).y, oPos(nextIdx + 1).z)
    curveElms.Add oLine3d(idx)
    Set oLine3d(idx) = Nothing
  Next idx

  
  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
  Set curveElms = Nothing
  
  
  oVec.Set xOffset, yOffset, 0
  oMat.LoadIdentity
  oMat.Translate oVec
  oCmplx.Transform oMat ' This positions the cutting shape with the current CP as origin. We assumed CP5 as origin while
  'creating the shape
  

  'modify the shape to take care of CopeLength,CopeDepth, Increment etc..
  ModifyCopeShape oSupping, oSupped, iEnd, True, oAttribsCAO, oCmplx
  
  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter
    
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub

Public Sub CreateBottomCopeforWShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
Const MT = "CreateBottomCopeforWShape"
  On Error GoTo ErrorHandler


  Dim depth As Double, tFlange As Double, tWeb As Double, bFlange As Double
  Dim cp As Long
  Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
  Dim oLine3d(1 To 4) As Line3d
  Dim pGeometryFactory As New GeometryFactory
  Dim oPos(1 To 4) As IJDPosition
  Dim idx As Integer
  Dim oCmplx As ComplexString3d
  Dim oMat As New DT4x4
  Dim oVec As New DVector
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim oArrayOfids(12) As Long
  Dim oSectionAttrbs As IJDAttributes
  Dim oProfileBO As ISPSCrossSection
  Dim curveElms As IJElements
  Dim WebClearance#, FlangeClearance1#, FlangeClearance2#
  
  WebClearance = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("WebClearance").Value
  FlangeClearance1 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance1").Value
  FlangeClearance2 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance2").Value
  
  Set oProfileBO = oSupping.CrossSection
  Set oSectionAttrbs = oSupping.CrossSection.definition
  
  depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
  bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
  tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
  tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value

  
  cp = oProfileBO.CardinalPoint

  oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
  oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.

  xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
  yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
  
  For idx = 1 To 4
    Set oPos(idx) = New DPosition
  Next idx
  
  'CP1 - oPos(1) CP3 - oPos(2)

  oPos(4).x = -1 * (bFlange / 2# + FlangeClearance1)
  oPos(4).y = -(depth / 2#) + (tFlange + FlangeClearance2)
  oPos(4).z = 0#
  
  oPos(3).x = -1 * oPos(4).x
  oPos(3).y = oPos(4).y
  oPos(3).z = 0#
  
  oPos(2).x = oPos(3).x
  oPos(2).y = -(depth / 2#) - FlangeClearance1
  oPos(2).z = 0#
  
  oPos(1).x = -1 * oPos(2).x
  oPos(1).y = oPos(2).y
  oPos(1).z = 0#
  
  ' Make edges away from the supported member longer so that we don't get any hanging pieces
  oPos(2).y = oPos(2).y - LARGE_EDGE
  oPos(1).y = oPos(1).y - LARGE_EDGE
  

  Set curveElms = New JObjectCollection
  For idx = 1 To 4
    Dim nextIdx As Integer
    nextIdx = idx Mod 4
    Set oLine3d(idx) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, _
    oPos(nextIdx + 1).x, oPos(nextIdx + 1).y, oPos(nextIdx + 1).z)
    
    curveElms.Add oLine3d(idx)
    Set oLine3d(idx) = Nothing
  Next idx

  
  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
  Set curveElms = Nothing
  
  
  oVec.Set xOffset, yOffset, 0
  oMat.LoadIdentity
  oMat.Translate oVec
  oCmplx.Transform oMat ' This positions the cutting shape with the current CP as origin. We assumed CP5 as origin while
  'creating the shape
  
  'modify the shape to take care of CopeLength,CopeDepth, Increment etc..
  ModifyCopeShape oSupping, oSupped, iEnd, False, oAttribsCAO, oCmplx
  
  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter
    
  For idx = 1 To 4
    Set oPos(idx) = Nothing
  Next idx
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub

Public Sub CreateTopCopeforCShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
Const MT = "CreateTopCopeforCShape"
  On Error GoTo ErrorHandler


  Dim depth As Double, tFlange As Double, tWeb As Double, bFlange As Double
  Dim cp As Long
  Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
  Dim oLine3d(1 To 4) As Line3d
  Dim pGeometryFactory As New GeometryFactory
  Dim oPos(1 To 4) As IJDPosition
  Dim idx As Integer
  Dim oCmplx As ComplexString3d
  Dim oMat As New DT4x4
  Dim oVec As New DVector
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim oArrayOfids(12) As Long
  Dim oSectionAttrbs As IJDAttributes
  Dim oProfileBO As ISPSCrossSection
  Dim WebClearance#, FlangeClearance1#, FlangeClearance2#

  
  WebClearance = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("WebClearance").Value
  FlangeClearance1 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance1").Value
  FlangeClearance2 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance2").Value
    
  Set oProfileBO = oSupping.CrossSection
  Set oSectionAttrbs = oSupping.CrossSection.definition
  
  depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
  bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
  tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
  tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value

  
  cp = oProfileBO.CardinalPoint

  oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
  oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.

  xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
  yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
  
  For idx = 1 To 4
    Set oPos(idx) = New DPosition
  Next idx
  
  'CP7 - oPos(1) CP9 - oPos(2)
  

  oPos(1).x = -1 * (bFlange / 2# + FlangeClearance1)
  oPos(1).y = (depth / 2#) + FlangeClearance1
  oPos(1).z = 0

  oPos(2).x = -1 * oPos(1).x
  oPos(2).y = oPos(1).y
  oPos(2).z = 0

  oPos(3).x = oPos(2).x
  oPos(3).y = (depth / 2#) - (tFlange + FlangeClearance2)
  oPos(3).z = 0

  oPos(4).x = -1 * oPos(3).x
  oPos(4).y = oPos(3).y
  oPos(4).z = 0
  ' Make edges away from the supported member longer so that we don't get any hanging pieces

  oPos(1).y = oPos(1).y + LARGE_EDGE
  oPos(2).y = oPos(2).y + LARGE_EDGE


  
  Dim curveElms As IJElements
  Set curveElms = New JObjectCollection
  
  For idx = 1 To 4
    Dim nextIdx As Integer
    nextIdx = idx Mod 4
    Set oLine3d(idx) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, _
    oPos(nextIdx + 1).x, oPos(nextIdx + 1).y, oPos(nextIdx + 1).z)
    
    curveElms.Add oLine3d(idx)
    Set oLine3d(idx) = Nothing
  Next idx

  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
  Set curveElms = Nothing



  oVec.Set xOffset, yOffset, 0
  
  oMat.LoadIdentity
  oMat.Translate oVec
  oCmplx.Transform oMat ' This positions the cutting shape with the current CP as origin. We assumed CP5 as origin while
  'creating the shape



  'modify the shape to take care of CopeLength,CopeDepth, Increment etc..
  ModifyCopeShape oSupping, oSupped, iEnd, True, oAttribsCAO, oCmplx

  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter
    
  For idx = 1 To 4
    Set oPos(idx) = Nothing
  Next idx
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub

Public Sub CreateBottomCopeforCShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, _
        iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
Const MT = "CreateBottomCopeforCShape"
  On Error GoTo ErrorHandler


  Dim depth As Double, tFlange As Double, tWeb As Double, bFlange As Double
  Dim cp As Long
  Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
  Dim oLine3d(1 To 4) As Line3d
  Dim pGeometryFactory As New GeometryFactory
  Dim oPos(1 To 4) As IJDPosition
  Dim idx As Integer
  Dim oCmplx As ComplexString3d
  Dim oMat As New DT4x4
  Dim oVec As New DVector
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim oArrayOfids(12) As Long
  Dim oSectionAttrbs As IJDAttributes
  Dim oProfileBO As ISPSCrossSection
  Dim curveElms As IJElements
  Dim WebClearance#, FlangeClearance1#, FlangeClearance2#
  
  
  WebClearance = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("WebClearance").Value
  FlangeClearance1 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance1").Value
  FlangeClearance2 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance2").Value
  
  Set oProfileBO = oSupping.CrossSection
  Set oSectionAttrbs = oSupping.CrossSection.definition
  
  depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
  bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
  tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
  tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value

  
  cp = oProfileBO.CardinalPoint

  oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
  oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.

  xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
  yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
  
  For idx = 1 To 4
    Set oPos(idx) = New DPosition
  Next idx
  
  'CP1 - oPos(1) CP3 - oPos(2)

  oPos(4).x = -1 * (bFlange / 2# + FlangeClearance1)
  oPos(4).y = -(depth / 2#) + (tFlange + FlangeClearance2)
  oPos(4).z = 0#
  
  oPos(3).x = -1 * oPos(4).x
  oPos(3).y = oPos(4).y
  oPos(3).z = 0#
  
  oPos(2).x = oPos(3).x
  oPos(2).y = -(depth / 2#) - FlangeClearance1
  oPos(2).z = 0#
  
  oPos(1).x = -1 * oPos(2).x
  oPos(1).y = oPos(2).y
  oPos(1).z = 0#
  
  ' Make edges away from the supported member longer so that we don't get any hanging pieces
  oPos(2).y = oPos(2).y - LARGE_EDGE
  oPos(1).y = oPos(1).y - LARGE_EDGE
  

  Set curveElms = New JObjectCollection
  For idx = 1 To 4
    Dim nextIdx As Integer
    nextIdx = idx Mod 4
    Set oLine3d(idx) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, _
    oPos(nextIdx + 1).x, oPos(nextIdx + 1).y, oPos(nextIdx + 1).z)
    
    curveElms.Add oLine3d(idx)
    Set oLine3d(idx) = Nothing
  Next idx

  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
  Set curveElms = Nothing


  oVec.Set xOffset, yOffset, 0

  oMat.LoadIdentity
  oMat.Translate oVec
  oCmplx.Transform oMat ' This positions the cutting shape with the current CP as origin. We assumed CP5 as origin while
  'creating the shape



  'modify the shape to take care of CopeLength,CopeDepth, Increment etc..
  ModifyCopeShape oSupping, oSupped, iEnd, False, oAttribsCAO, oCmplx

  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter

  For idx = 1 To 4
    Set oPos(idx) = Nothing
  Next idx
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub


Public Sub CreateBottomCopeforLShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
Const MT = "CreateBottomCopeforLShape"
  On Error GoTo ErrorHandler


  Dim depth As Double, tFlange As Double, bFlange As Double
  Dim cp As Long
  Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
  Dim oLine3d(1 To 4) As Line3d
  Dim pGeometryFactory As New GeometryFactory
  Dim oPos(1 To 4) As IJDPosition
  Dim idx As Integer
  Dim oCmplx As ComplexString3d
  Dim oMat As New DT4x4
  Dim oVec As New DVector
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim oArrayOfids(12) As Long
  Dim oSectionAttrbs As IJDAttributes
  Dim oProfileBO As ISPSCrossSection
  Dim curveElms As IJElements
  Dim iQuadrant As Integer
  Set oProfileBO = oSupping.CrossSection
  Dim FlangeClearance1#, FlangeClearance2#
  

  FlangeClearance1 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance1").Value
  FlangeClearance2 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance2").Value

  Set oSectionAttrbs = oSupping.CrossSection.definition
  
  depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
  bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
  tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value

  
  cp = oProfileBO.CardinalPoint

  oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
  oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.

  xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
  yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
  
  For idx = 1 To 4
    Set oPos(idx) = New DPosition
  Next idx
  
  iQuadrant = GetIncidentMemberQuadrant(oSupping, oSupped, iEnd)
  'we shouldn't be creating cope if iQuadrant <> 1 or 2
  If iQuadrant = 1 Then
      '  oPos(1)-CP7 , oPos(4) - CP1  ,clockwise from oPos(1)
      oPos(1).x = -(bFlange / 2# + FlangeClearance1)
      oPos(1).y = (depth / 2#) + FlangeClearance1
      oPos(1).z = 0#
      
      oPos(2).x = -(bFlange / 2# - tFlange - FlangeClearance2)
      oPos(2).y = oPos(1).y
      oPos(2).z = 0#
      
      oPos(3).x = oPos(2).x
      oPos(3).y = -oPos(2).y
      oPos(3).z = 0#
      
      oPos(4).x = oPos(1).x
      oPos(4).y = oPos(3).y
      oPos(4).z = 0#
      
      ' Make edges away from the supported member longer so that we don't get any hanging pieces
      oPos(3).y = oPos(3).y - LARGE_EDGE
      oPos(4).y = oPos(4).y - LARGE_EDGE
  ElseIf (iQuadrant = 2) Or (iQuadrant = 4) Then
      'CP1 - oPos(1) CP3 - oPos(2)
      'we create this points based on iQuadrant =2. The transform will take care
      'if supported is actually in iQuadrant =4
      oPos(4).x = -1 * (bFlange / 2# + FlangeClearance1)
      oPos(4).y = -(depth / 2#) + (tFlange + FlangeClearance2)
      oPos(4).z = 0#
      
      oPos(3).x = -1 * oPos(4).x
      oPos(3).y = oPos(4).y
      oPos(3).z = 0#
      
      oPos(2).x = oPos(3).x
      oPos(2).y = -(depth / 2#) - FlangeClearance1
      oPos(2).z = 0#
      
      oPos(1).x = oPos(4).x
      oPos(1).y = oPos(2).y
      oPos(1).z = 0#
      
      ' Make edges away from the supported member longer so that we don't get any hanging pieces
      oPos(2).y = oPos(2).y - LARGE_EDGE
      oPos(1).y = oPos(1).y - LARGE_EDGE
  End If
  

  Set curveElms = New JObjectCollection
  For idx = 1 To 4
    Dim nextIdx As Integer
    nextIdx = idx Mod 4
    Set oLine3d(idx) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, _
    oPos(nextIdx + 1).x, oPos(nextIdx + 1).y, oPos(nextIdx + 1).z)
    
    curveElms.Add oLine3d(idx)
    Set oLine3d(idx) = Nothing
  Next idx

  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
  Set curveElms = Nothing

  oVec.Set xOffset, yOffset, 0

  oMat.LoadIdentity
  oMat.Translate oVec
  oCmplx.Transform oMat ' This positions the cutting shape with the current CP as origin. We assumed CP5 as origin while
  'creating the shape

    
  'modify the shape to take care of CopeLength,CopeDepth, Increment etc..

  If (iQuadrant = 2) Or (iQuadrant = 4) Then
    ModifyCopeShape oSupping, oSupped, iEnd, False, oAttribsCAO, oCmplx
  End If

  
  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter
    
  For idx = 1 To 4
    Set oPos(idx) = Nothing
  Next idx
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub

Public Sub CreateTopCopeforTShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
Const MT = "CreateTopCopeforTShape"
  On Error GoTo ErrorHandler


  Dim depth As Double, tFlange As Double, tWeb As Double, bFlange As Double
  Dim cp As Long
  Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
  Dim oLine3d(1 To 4) As Line3d
  Dim pGeometryFactory As New GeometryFactory
  Dim oPos(1 To 4) As IJDPosition
  Dim idx As Integer
  Dim oCmplx As ComplexString3d
  Dim oMat As New DT4x4
  Dim oVec As New DVector
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim oArrayOfids(12) As Long
  Dim oSectionAttrbs As IJDAttributes
  Dim oProfileBO As ISPSCrossSection
  Dim iQuadrant As Integer
  Dim WebClearance#, FlangeClearance1#, FlangeClearance2#
    
  WebClearance = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("WebClearance").Value
  FlangeClearance1 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance1").Value
  FlangeClearance2 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance2").Value
    
  Set oProfileBO = oSupping.CrossSection
  Set oSectionAttrbs = oSupping.CrossSection.definition
  
  depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
  bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
  tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
  tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value

  
  cp = oProfileBO.CardinalPoint

  oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
  oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.

  xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
  yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
  
  For idx = 1 To 4
    Set oPos(idx) = New DPosition
  Next idx
  
  iQuadrant = GetIncidentMemberQuadrant(oSupping, oSupped, iEnd)
  If iQuadrant = 2 Or iQuadrant = 4 Then
      'CP7 - oPos(1) CP9 - oPos(2)
      
    
      oPos(1).x = -1 * (bFlange / 2# + FlangeClearance1)
      oPos(1).y = (depth / 2#) + FlangeClearance1
      oPos(1).z = 0
    
      oPos(2).x = -1 * oPos(1).x
      oPos(2).y = oPos(1).y
      oPos(2).z = 0
    
      oPos(3).x = oPos(2).x
      oPos(3).y = (depth / 2#) - (tFlange + FlangeClearance2)
      oPos(3).z = 0
    
      oPos(4).x = -1 * oPos(3).x
      oPos(4).y = oPos(3).y
      oPos(4).z = 0
      ' Make edges away from the supported member longer so that we don't get any hanging pieces
      oPos(1).y = oPos(1).y + LARGE_EDGE
      oPos(2).y = oPos(2).y + LARGE_EDGE
      
  ElseIf iQuadrant = 3 Then

      oPos(1).x = -1 * (tWeb / 2# + FlangeClearance1)
      oPos(1).y = (depth / 2#) + FlangeClearance1
      oPos(1).z = 0
    
      oPos(2).x = -1 * oPos(1).x
      oPos(2).y = oPos(1).y
      oPos(2).z = 0
    
      oPos(3).x = oPos(2).x
      oPos(3).y = -oPos(2).y
      oPos(3).z = 0
    
      oPos(4).x = oPos(1).x
      oPos(4).y = oPos(3).y
      oPos(4).z = 0
      ' Make edges away from the supported member longer so that we don't get any hanging pieces
      oPos(1).y = oPos(1).y + LARGE_EDGE
      oPos(2).y = oPos(2).y + LARGE_EDGE
  End If
  

  
  Dim curveElms As IJElements
  Set curveElms = New JObjectCollection
  
  For idx = 1 To 4
    Dim nextIdx As Integer
    nextIdx = idx Mod 4
    Set oLine3d(idx) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, _
    oPos(nextIdx + 1).x, oPos(nextIdx + 1).y, oPos(nextIdx + 1).z)
    
    curveElms.Add oLine3d(idx)
    Set oLine3d(idx) = Nothing
  Next idx

  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
  Set curveElms = Nothing


  oVec.Set xOffset, yOffset, 0
  
  oMat.LoadIdentity
  oMat.Translate oVec
  oCmplx.Transform oMat ' This positions the cutting shape with the current CP as origin. We assumed CP5 as origin while
  'creating the shape



  'modify the shape to take care of CopeLength,CopeDepth, Increment etc..
  'don't know yet how to do this if iQuadrant=3
  If iQuadrant = 2 Or iQuadrant = 4 Then
    ModifyCopeShape oSupping, oSupped, iEnd, True, oAttribsCAO, oCmplx
  End If

  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter
    
  For idx = 1 To 4
    Set oPos(idx) = Nothing
  Next idx
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub
Public Sub CreateBottomCopefor2LShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, _
        iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
Const MT = "CreateBottomCopefor2LShape"
  On Error GoTo ErrorHandler


  Dim depth As Double, tFlange As Double, sBackToBack As Double, bFlange As Double
  Dim cp As Long
  Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
  Dim oLine3d(1 To 4) As Line3d
  Dim pGeometryFactory As New GeometryFactory
  Dim oPos(1 To 4) As IJDPosition
  Dim idx As Integer
  Dim oCmplx As ComplexString3d
  Dim oMat As New DT4x4
  Dim oVec As New DVector
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim oArrayOfids(12) As Long
  Dim oSectionAttrbs As IJDAttributes
  Dim oProfileBO As ISPSCrossSection
  Dim iQuadrant As Integer
  Dim WebClearance#, FlangeClearance1#, FlangeClearance2#
  
  WebClearance = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("WebClearance").Value
  FlangeClearance1 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance1").Value
  FlangeClearance2 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance2").Value
    
  Set oProfileBO = oSupping.CrossSection
  Set oSectionAttrbs = oSupping.CrossSection.definition
  
  depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
  bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
  tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
  sBackToBack = oSectionAttrbs.CollectionOfAttributes("IJUA2L").Item("bb").Value
  
  
  cp = oProfileBO.CardinalPoint

  oProfileBO.GetCardinalPointOffset cp, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP which is our -y and z.
  oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of CP=5 which is our -y and z.

  xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
  yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
  
  For idx = 1 To 4
    Set oPos(idx) = New DPosition
  Next idx
  
  iQuadrant = GetIncidentMemberQuadrant(oSupping, oSupped, iEnd)
  If iQuadrant = 2 Or iQuadrant = 4 Then
      'CP1 - oPos(1) CP3 - oPos(2)
      oPos(4).x = -1 * (bFlange / 2# + FlangeClearance1)
      oPos(4).y = -(depth / 2#) + tFlange + FlangeClearance2
      oPos(4).z = 0
    
      oPos(3).x = -1 * oPos(4).x
      oPos(3).y = oPos(4).y
      oPos(3).z = 0
    
      oPos(2).x = oPos(3).x
      oPos(2).y = -(depth / 2#) - FlangeClearance1
      oPos(2).z = 0
    
      oPos(1).x = -1 * oPos(2).x
      oPos(1).y = oPos(2).y
      oPos(1).z = 0
      ' Make edges away from the supported member longer so that we don't get any hanging pieces
      oPos(2).y = oPos(2).y - LARGE_EDGE
      oPos(1).y = oPos(1).y - LARGE_EDGE
  ElseIf iQuadrant = 1 Then

      oPos(1).x = -1 * (tFlange + sBackToBack / 2 + FlangeClearance1)
      oPos(1).y = (depth / 2#) + FlangeClearance1
      oPos(1).z = 0
    
      oPos(2).x = -1 * oPos(1).x
      oPos(2).y = oPos(1).y
      oPos(2).z = 0
    
      oPos(3).x = oPos(2).x
      oPos(3).y = -oPos(2).y
      oPos(3).z = 0
    
      oPos(4).x = oPos(1).x
      oPos(4).y = oPos(3).y
      oPos(4).z = 0
      ' Make edges away from the supported member longer so that we don't get any hanging pieces
      oPos(3).y = oPos(3).y - LARGE_EDGE
      oPos(4).y = oPos(4).y - LARGE_EDGE
  End If
 
  Dim curveElms As IJElements
  Set curveElms = New JObjectCollection
  
  For idx = 1 To 4
    Dim nextIdx As Integer
    nextIdx = idx Mod 4
    Set oLine3d(idx) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, _
    oPos(nextIdx + 1).x, oPos(nextIdx + 1).y, oPos(nextIdx + 1).z)
    
    curveElms.Add oLine3d(idx)
    Set oLine3d(idx) = Nothing
  Next idx

  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
  Set curveElms = Nothing
  
  oVec.Set xOffset, yOffset, 0
  
  oMat.LoadIdentity
  oMat.Translate oVec
  oCmplx.Transform oMat ' This positions the cutting shape with the current CP as origin. We assumed CP5 as origin while
  'creating the shape


  'modify the shape to take care of CopeLength,CopeDepth, Increment etc..
  'don't know yet how to do this if iQuadrant<>2
  If iQuadrant = 2 Or iQuadrant = 4 Then
    ModifyCopeShape oSupping, oSupped, iEnd, False, oAttribsCAO, oCmplx
  End If

  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter
  
  For idx = 1 To 4
    Set oPos(idx) = Nothing
  Next idx
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub
Public Sub CreateCutback(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oPlane As IJPlane)
Const MT = "CreateCutback"
    On Error GoTo ErrorHandler
    
    Dim iQuadrant As Integer
    Dim pGeometryFactory As New GeometryFactory
    Dim WebClearance#, FlangeClearance#
    Dim oAttribsColl As IJDAttributesCol
    Dim oPos As IJDPosition, oPosAlongAxis As IJDPosition
    Dim oVec As IJDVector
    Dim Clearance#
    
    iQuadrant = GetIncidentMemberQuadrant(oSupping, oSupped, iEnd)
    On Error Resume Next
        Set oAttribsColl = oAttribsCAO.CollectionOfAttributes("IJUASPSCope")
    On Error GoTo ErrorHandler
    If Not oAttribsColl Is Nothing Then
        WebClearance = oAttribsColl.Item("WebClearance").Value
        FlangeClearance = oAttribsColl.Item("FlangeClearance1").Value
        If (iQuadrant = 1) Or (iQuadrant = 3) Then
            'use flange clearnce, if supped is connected to the flange
            Clearance = FlangeClearance
        Else
            'use web clearnce, if supped is connected to the web
            Clearance = WebClearance
        End If
    Else
        Set oAttribsColl = oAttribsCAO.CollectionOfAttributes("IJUASPSPlanarCutback")
        Clearance = oAttribsColl.Item("Clearance").Value
    End If
    
    Set oPosAlongAxis = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)

    Set oPos = GetMembSidePos(oSupping, iQuadrant, oPosAlongAxis)
    Set oVec = GetMembSidePlaneNormal(oSupping, iQuadrant, oPosAlongAxis)
    
    oPos.x = oPos.x + oVec.x * Clearance
    oPos.y = oPos.y + oVec.y * Clearance
    oPos.z = oPos.z + oVec.z * Clearance
    
    If oPlane Is Nothing Then
        Set oPlane = pGeometryFactory.Planes3d.CreateByPointNormal(oRescMgr, oPos.x, oPos.y, oPos.z, oVec.x, oVec.y, oVec.z)
    Else
        oPlane.SetRootPoint oPos.x, oPos.y, oPos.z
        oPlane.SetNormal oVec.x, oVec.y, oVec.z
    End If
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub

Public Sub CreateCutbackForRectShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oPlane As IJPlane)
Const MT = "CreateCutbackForRectShape"
    On Error GoTo ErrorHandler
    
    Dim iQuadrant As Integer
    Dim pGeometryFactory As New GeometryFactory
    Dim WebClearance#, FlangeClearance#
    Dim oAttribsColl As IJDAttributesCol
    Dim oPos As IJDPosition, oPosAlongAxis As IJDPosition
    Dim oVec As IJDVector
    
    On Error Resume Next
        Set oAttribsColl = oAttribsCAO.CollectionOfAttributes("IJUASPSCope")
    On Error GoTo ErrorHandler
    If Not oAttribsColl Is Nothing Then
        WebClearance = oAttribsColl.Item("WebClearance").Value
        FlangeClearance = oAttribsColl.Item("FlangeClearance1").Value
    Else
        Set oAttribsColl = oAttribsCAO.CollectionOfAttributes("IJUASPSPlanarCutback")
        WebClearance = oAttribsColl.Item("Clearance").Value
        FlangeClearance = oAttribsColl.Item("Clearance").Value
    End If
    
    
    iQuadrant = GetIncidentMemberQuadrant(oSupping, oSupped, iEnd)
    Set oPosAlongAxis = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)
    Set oPos = GetSectionRangeBoxSidePos(oSupping, iQuadrant, oPosAlongAxis)
    Set oVec = GetMembSidePlaneNormal(oSupping, iQuadrant, oPosAlongAxis)
    
    oPos.x = oPos.x + oVec.x * FlangeClearance
    oPos.y = oPos.y + oVec.y * FlangeClearance
    oPos.z = oPos.z + oVec.z * FlangeClearance
    
    If oPlane Is Nothing Then
        Set oPlane = pGeometryFactory.Planes3d.CreateByPointNormal(oRescMgr, oPos.x, oPos.y, oPos.z, oVec.x, oVec.y, oVec.z)
    Else
        oPlane.SetRootPoint oPos.x, oPos.y, oPos.z
        oPlane.SetNormal oVec.x, oVec.y, oVec.z
    End If
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub
Public Sub CreateSquaredCutbackForRectShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oPlane As IJPlane)
Const MT = "CreateSquaredCutbackForRectShape"
    On Error GoTo ErrorHandler
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#, dist#
    Dim oPos As IJDPosition
    Dim oMat As IJDT4x4
    Dim oVec As New DVector
    Dim oAxisCurve As IJCurve
    Dim oSurface As IJSurface
    
    
    oSupped.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    'get the vector towards the other end
    If iEnd = SPSMemberAxisStart Then
        oSupped.Rotation.GetTransformAtPosition sX, sY, sZ, oMat, Nothing
        oVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    Else
        oSupped.Rotation.GetTransformAtPosition eX, eY, eZ, oMat, Nothing
        oVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
        oVec.[Scale] -1 'reverse direction
    End If
    

    CreateCutbackForRectShape oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oPlane
    'from the supping axis and plane, create a surface of projection or revolution
    CreateCutbackSurface Nothing, oSupping, oSupped, iEnd, oPlane, oSurface
    
    GetCutDistAndPosFromPlane oSupped, iEnd, oSurface, oPos, dist
    oPlane.SetNormal oVec.x, oVec.y, oVec.z
    oPlane.SetRootPoint oPos.x, oPos.y, oPos.z
    
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub


Public Function GetMirrorMatrix() As IJDT4x4
Const MT = "GetMirrorMatrix"
  On Error GoTo ErrorHandler
    Dim pMat(15) As Double
    Dim oMat As IJDT4x4
    Set oMat = New DT4x4
    oMat.LoadIdentity
    oMat.Get pMat(0)
    
    pMat(0) = -1#
    pMat(1) = 0#
    pMat(2) = 0#
    pMat(4) = 0#
    pMat(5) = 1 '-1#
    pMat(6) = 0#
    pMat(8) = 0#
    pMat(9) = 0#
    pMat(10) = 1#
    oMat.Set pMat(0)
    Set GetMirrorMatrix = oMat
    Set oMat = Nothing
  Exit Function
ErrorHandler:    HandleError MODULE, MT
End Function


Public Sub CreateSquaredCutbackForWShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oCutbackPlane As IJPlane)
Const MT = "CreateSquaredCutbackForWShape"
    On Error GoTo ErrorHandler
    
    Dim oCornerLines() As IJLine
    Dim oInterSecPoints() As IJDPosition
    Dim i As Integer
    Dim lBnd As Integer, uBnd As Integer
    Dim oPos As New DPosition
    Dim x#, y#, z#, nx#, ny#, nz#
    Dim oMat As IJDT4x4
    Dim oSectionAttrbs As IJDAttributes
    Dim oLeftBottomOfTopFlange As New DPosition, oRightBottomOfTopFlange As New DPosition
    Dim oLeftTopOfBottomFlange As New DPosition, oRightTopOfBottomFlange As New DPosition
    Dim oLeftBottomOfWeb As New DPosition, oRightBottomOfWeb As New DPosition
    Dim oLeftTopOfWeb As New DPosition, oRightTopOfWeb As New DPosition
    Dim FlangeClearance1#, FlangeClearance2#, WebClearance#
    Dim depth As Double, tFlange As Double, tWeb As Double, bFlange As Double
    Dim oFlangeSidePlane As IJPlane
    Dim pGeometryFactory As New GeometryFactory
    Dim oVec As IJDVector
    Dim bIsSuppedAxisInPositiveY As Boolean, bIsSuppedAxisInPositiveZ As Boolean
    Dim oLine(0) As IJLine
    Dim oSurf As IJSurface
    Dim code As Geom3dIntersectConstants
    Dim colElms As IJElements
    Dim oRootPoint As IJPoint
    Dim oClippedPos() As IJDPosition
    Dim oBottomClipPos As New DPosition
    Dim oTopClipPos As New DPosition
    Dim oPosAlongAxis As IJDPosition
    Dim oAxisCurve As IJCurve
    Dim oSurface As IJSurface
    
    Set oSectionAttrbs = oSupping.CrossSection.definition
    WebClearance = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("WebClearance").Value
    FlangeClearance1 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance1").Value
    FlangeClearance2 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance2").Value
    depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
    tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
    
    'get the cutback plane which is not squared yet
    CreateCutback oSupping, oSupped, iEnd, oAttribsCAO, oRescMgr, oCutbackPlane
    
    'from the supping axis and plane, create a surface of projection or revolution
    CreateCutbackSurface Nothing, oSupping, oSupped, iEnd, oCutbackPlane, oSurface
    
    Set oPosAlongAxis = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)
    'set flange points
    oLeftBottomOfTopFlange.Set -(bFlange / 2 + FlangeClearance1), (depth / 2 - tFlange - FlangeClearance2), 0
    oRightBottomOfTopFlange.Set (bFlange / 2 + FlangeClearance1), (depth / 2 - tFlange - FlangeClearance2), 0
    oLeftTopOfBottomFlange.Set -(bFlange / 2 + FlangeClearance1), -(depth / 2 - tFlange - FlangeClearance2), 0
    oRightTopOfBottomFlange.Set (bFlange / 2 + FlangeClearance1), -(depth / 2 - tFlange - FlangeClearance2), 0
    
    'set web points
    oLeftTopOfWeb.Set -(tWeb / 2 + WebClearance), (depth / 2 - tFlange - FlangeClearance2), 0
    oLeftBottomOfWeb.Set -(tWeb / 2 + WebClearance), -(depth / 2 - tFlange - FlangeClearance2), 0
    oRightTopOfWeb.Set (tWeb / 2 + WebClearance), (depth / 2 - tFlange - FlangeClearance2), 0
    oRightBottomOfWeb.Set (tWeb / 2 + WebClearance), -(depth / 2 - tFlange - FlangeClearance2), 0

    'transform Flange and web points to global
    Set oLeftBottomOfTopFlange = TransformPosToGlobal(oSupping, oLeftBottomOfTopFlange, oPosAlongAxis)
    Set oRightBottomOfTopFlange = TransformPosToGlobal(oSupping, oRightBottomOfTopFlange, oPosAlongAxis)
    Set oLeftTopOfBottomFlange = TransformPosToGlobal(oSupping, oLeftTopOfBottomFlange, oPosAlongAxis)
    Set oRightTopOfBottomFlange = TransformPosToGlobal(oSupping, oRightTopOfBottomFlange, oPosAlongAxis)
    
    Set oLeftTopOfWeb = TransformPosToGlobal(oSupping, oLeftTopOfWeb, oPosAlongAxis)
    Set oLeftBottomOfWeb = TransformPosToGlobal(oSupping, oLeftBottomOfWeb, oPosAlongAxis)
    Set oRightTopOfWeb = TransformPosToGlobal(oSupping, oRightTopOfWeb, oPosAlongAxis)
    Set oRightBottomOfWeb = TransformPosToGlobal(oSupping, oRightBottomOfWeb, oPosAlongAxis)
    
    
    If IsSupportedAxisInPositiveY(oSupping, oSupped, iEnd) = True Then
        bIsSuppedAxisInPositiveY = True
    End If
    If IsSupportedAxisInPositiveZ(oSupping, oSupped, iEnd) = True Then
        bIsSuppedAxisInPositiveZ = True
    End If
    
    'get lines connecting cross section corner points
    GetCrossSectionLines oSupped, iEnd, oCornerLines
    
    If bIsSuppedAxisInPositiveY Then
        If oSupping.Rotation.Mirror = False Then
            If bIsSuppedAxisInPositiveZ Then
                Set oTopClipPos = oLeftBottomOfTopFlange
                Set oBottomClipPos = oLeftBottomOfWeb
            Else
                Set oTopClipPos = oLeftTopOfBottomFlange
                Set oBottomClipPos = oLeftTopOfWeb
            End If
        Else
            If bIsSuppedAxisInPositiveZ Then
                Set oTopClipPos = oRightBottomOfTopFlange
                Set oBottomClipPos = oRightBottomOfWeb
            Else
                Set oTopClipPos = oRightTopOfBottomFlange
                Set oBottomClipPos = oRightTopOfWeb
            End If
        End If
    Else
        If oSupping.Rotation.Mirror = False Then
            If bIsSuppedAxisInPositiveZ Then
                Set oTopClipPos = oRightBottomOfTopFlange
                Set oBottomClipPos = oRightBottomOfWeb
            Else
                Set oTopClipPos = oRightTopOfBottomFlange
                Set oBottomClipPos = oRightTopOfWeb
            End If
        Else
            If bIsSuppedAxisInPositiveZ Then
                Set oTopClipPos = oLeftBottomOfTopFlange
                Set oBottomClipPos = oLeftBottomOfWeb
            Else
                Set oTopClipPos = oLeftTopOfBottomFlange
                Set oBottomClipPos = oLeftTopOfWeb
            End If
        End If
    End If
    
    'todo-function need to take oPosalongAxis as a parameter
    Set oMat = GetCopeTransform(oSupping, oSupped, iEnd)
    
    oSupped.PointAtEnd(iEnd).GetPoint x, y, z
    oPosAlongAxis.Set x, y, z
 
    Set oVec = GetXVector(oSupped, oPosAlongAxis)
    oVec.Get nx, ny, nz
    oTopClipPos.Get x, y, z
    Set oLine(0) = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, x, y, z, nx, ny, nz, 1#)
    'now intersect this line with the cutback plane which is not yet squared
    GetIntersectionPoints oLine, oSurface, oInterSecPoints

    lBnd = LBound(oInterSecPoints)
    Set oTopClipPos = oInterSecPoints(lBnd)

    
    'now intersect the cross section lines with the cutback plane plane which is not yet squared
    GetIntersectionPoints oCornerLines, oSurface, oInterSecPoints
    
    'now clip the positions which are removed due to the cope
    ClipEdgePosToYDepth oInterSecPoints, oTopClipPos, oBottomClipPos, oMat, oClippedPos
    
    
    GetMembLineFromEndPos oSupped, iEnd, oClippedPos, oCornerLines
    'get intersection point that is closest to the other end
    Set oPos = GetPosClosestToOtherEnd(oClippedPos, oCornerLines, iEnd)
    Set oVec = GetXVector(oSupped, oPosAlongAxis)
    ' now square the cutback plane
    oCutbackPlane.SetRootPoint oPos.x, oPos.y, oPos.z
    oCutbackPlane.SetNormal oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2) 'oVec.x, oVec.y, oVec.z
    
    Set oLine(0) = Nothing
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub


Public Sub CreateSquaredCopeForWShape(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oAttribsCAO As IJDAttributes, bTop As Boolean, oRescMgr As IUnknown, ByRef oCutter As IJDModelBody)
Const MT = "CreateSquaredCopeForWShape"
  On Error GoTo ErrorHandler
    Dim oCornerLines() As IJLine
    Dim oInterSecPoints() As IJDPosition
    Dim i As Integer
    Dim lBnd As Integer, uBnd As Integer
    Dim oPos As New DPosition
    Dim x#, y#, z#, nx#, ny#, nz#
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim oMat As IJDT4x4
    Dim oSectionAttrbs As IJDAttributes
    Dim oLeftBottomOfTopFlange As New DPosition, oRightBottomOfTopFlange As New DPosition
    Dim oLeftTopOfBottomFlange As New DPosition, oRightTopOfBottomFlange As New DPosition
    Dim FlangeClearance1#, FlangeClearance2#, WebClearance#
    Dim depth As Double, tFlange As Double, tWeb As Double, bFlange As Double
    Dim oFlangeSidePlane As IJPlane
    Dim pGeometryFactory As New GeometryFactory
    Dim oVec As IJDVector
    Dim bIsSuppedAxisInPositiveY As Boolean, bIsSuppedAxisInPositiveZ As Boolean
    Dim oLine As IJLine
    Dim oSurf As IJSurface
    Dim code As Geom3dIntersectConstants
    Dim colElms As IJElements
    Dim oIntrSecPoint As IJPoint
    Dim oMembObjetcs As IJDMemberObjects
    Dim oCopePlane As IJPlane
    Dim oTopCopePos As New DPosition, oBottCopePos As New DPosition
    Dim oSuppingZvec As IJDVector
    Dim oCopeYvec As IJDVector, oCopeXvec As New DVector, oCopeZvec As IJDVector
    Dim oCutbackPlane As IJPlane
    Dim oCopePos(1 To 4) As IJDPosition
    Dim intFlag As Integer
    Dim oCopeWebPos As New DPosition
    Dim dblCopedepth#, dblCopeLength#
    Dim dblDist#
    Dim oIntersecPos As New DPosition
    Dim oPosAlongAxis As IJDPosition
    Dim oAxisCurve As IJCurve
    Dim oSurface As IJSurface
    
    
    Set oSectionAttrbs = oSupping.CrossSection.definition
    WebClearance = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("WebClearance").Value
    FlangeClearance1 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance1").Value
    FlangeClearance2 = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("FlangeClearance2").Value
    dblCopedepth = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("Depth").Value
    dblCopeLength = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("Length").Value
    
    
    depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
    tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
    
    'we need to get the cutback plane from the custom assembly
    Set oMembObjetcs = oAttribsCAO
    Set oCutbackPlane = oMembObjetcs.Item(1) 'cutback is the first member of the assembly
    
    
    
    Set oPosAlongAxis = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)
    
    'set flange points
    oLeftBottomOfTopFlange.Set -(bFlange / 2 + FlangeClearance1), (depth / 2 - tFlange - FlangeClearance2), 0
    oRightBottomOfTopFlange.Set (bFlange / 2 + FlangeClearance1), (depth / 2 - tFlange - FlangeClearance2), 0
    oLeftTopOfBottomFlange.Set -(bFlange / 2 + FlangeClearance1), -(depth / 2 - tFlange - FlangeClearance2), 0
    oRightTopOfBottomFlange.Set (bFlange / 2 + FlangeClearance1), -(depth / 2 - tFlange - FlangeClearance2), 0
    

    'Flange points are  global now
    Set oLeftBottomOfTopFlange = TransformPosToGlobal(oSupping, oLeftBottomOfTopFlange, oPosAlongAxis)
    Set oRightBottomOfTopFlange = TransformPosToGlobal(oSupping, oRightBottomOfTopFlange, oPosAlongAxis)
    Set oLeftTopOfBottomFlange = TransformPosToGlobal(oSupping, oLeftTopOfBottomFlange, oPosAlongAxis)
    Set oRightTopOfBottomFlange = TransformPosToGlobal(oSupping, oRightTopOfBottomFlange, oPosAlongAxis)
    
    ' now create a plane with either left or right flnage point as the root point and supping Y axis as normal
    'This plane represents the bounding plane  either the left or right side.
    If IsSupportedAxisInPositiveY(oSupping, oSupped, iEnd) = True Then
        bIsSuppedAxisInPositiveY = True
    End If
    If IsSupportedAxisInPositiveZ(oSupping, oSupped, iEnd) = True Then
        bIsSuppedAxisInPositiveZ = True
    End If
    If bIsSuppedAxisInPositiveY Then
        If oSupping.Rotation.Mirror = False Then
            oLeftBottomOfTopFlange.Get x, y, z
        Else
            oRightBottomOfTopFlange.Get x, y, z
        End If
    Else
        If oSupping.Rotation.Mirror = False Then
            oRightBottomOfTopFlange.Get x, y, z
        Else
            oLeftBottomOfTopFlange.Get x, y, z
        End If
    End If
    
    Set oVec = GetYVector(oSupping, oPosAlongAxis)
    oVec.Get nx, ny, nz
    Set oFlangeSidePlane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, x, y, z, nx, ny, nz)
    
    
    'from the supping axis and plane, create a surface of projection or revolution
    CreateCutbackSurface Nothing, oSupping, oSupped, iEnd, oFlangeSidePlane, oSurface
    
  
    'get lines connecting cross section corner points
    GetCrossSectionLines oSupped, iEnd, oCornerLines
    
    'now intersect the cross section lines with the flange side  plane
    GetIntersectionPoints oCornerLines, oSurface, oInterSecPoints
    
    'get intersection point that is farthest from the end being connected.
    'that is closest to other end
    Set oIntersecPos = GetPosClosestToOtherEnd(oInterSecPoints, oCornerLines, iEnd)

    oIntersecPos.Get x, y, z
    'get supped axis direction
    oSupped.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    If iEnd = SPSMemberAxisStart Then
        oSupped.Rotation.GetTransformAtPosition sX, sY, sZ, oMat, Nothing
        oCopeXvec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    Else
        oSupped.Rotation.GetTransformAtPosition eX, eY, eZ, oMat, Nothing
        oCopeXvec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
        oCopeXvec.[Scale] -1 'reverse direction such that the vector is away from the connected end
    End If

    oCopeXvec.length = 1
    
    Set oSuppingZvec = GetZVector(oSupping, oPosAlongAxis)
    Set oCopeZvec = oCopeXvec.Cross(oSuppingZvec)
    oCopeZvec.length = 1
    Set oCopeYvec = oCopeZvec.Cross(oCopeXvec)
    oCopeYvec.length = 1
    
    oCopeXvec.Get nx, ny, nz
    
    'create plane with the intersection point(that gives max cope distance) as root point and supped x axis as the normal
    'this plane is parallel to the cutback plane but offset by the cope length
    Set oCopePlane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, x, y, z, nx, ny, nz)
    
    oSupping.Rotation.GetTransformAtPosition oPosAlongAxis.x, oPosAlongAxis.y, oPosAlongAxis.z, oMat, Nothing
    oMat.Invert 'global to local
    'transform intersection point to supping local coordinates
    Set oPos = oMat.TransformPosition(oIntersecPos)
    oPos.Get x, y, z
    
    
    'transform flange point to supping local coordinates
    Set oLeftBottomOfTopFlange = oMat.TransformPosition(oLeftBottomOfTopFlange)
    Set oLeftTopOfBottomFlange = oMat.TransformPosition(oLeftTopOfBottomFlange)
    Set oRightBottomOfTopFlange = oMat.TransformPosition(oRightBottomOfTopFlange)
    Set oRightTopOfBottomFlange = oMat.TransformPosition(oRightTopOfBottomFlange)
    
    'change position of flange points such that they are at the same local x as the intersection point
    oLeftBottomOfTopFlange.x = x
    oLeftTopOfBottomFlange.x = x
    oRightBottomOfTopFlange.x = x
    oRightTopOfBottomFlange.x = x
    oMat.Invert 'local to global
    Set oLeftBottomOfTopFlange = oMat.TransformPosition(oLeftBottomOfTopFlange)
    Set oLeftTopOfBottomFlange = oMat.TransformPosition(oLeftTopOfBottomFlange)
    Set oRightBottomOfTopFlange = oMat.TransformPosition(oRightBottomOfTopFlange)
    Set oRightTopOfBottomFlange = oMat.TransformPosition(oRightTopOfBottomFlange)
    
    If bIsSuppedAxisInPositiveY Then
        'create another point which is projection of flange point along supped x axis
        If bIsSuppedAxisInPositiveZ Then
            oLeftBottomOfTopFlange.Get x, y, z
        Else
            oLeftTopOfBottomFlange.Get x, y, z
        End If
    Else
        If bIsSuppedAxisInPositiveZ Then
            oRightBottomOfTopFlange.Get x, y, z
        Else
            oRightTopOfBottomFlange.Get x, y, z
        End If
    End If
    'create line through flange point parallel to supped axis
    Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, x, y, z, nx, ny, nz, 1#)
    'make the line infinite
    oLine.Infinite = True
    Set oSurf = oCopePlane
    oSurf.Intersect oLine, colElms, code
    If Not colElms Is Nothing Then
        If colElms.count > 0 Then
            Set oIntrSecPoint = colElms.Item(1)
            oIntrSecPoint.GetPoint x, y, z
            oPos.Set x, y, z
            GetCopePoint oSupping, oSupped, iEnd, oCopePlane, bIsSuppedAxisInPositiveZ, oIntrSecPoint
            oIntrSecPoint.GetPoint x, y, z
            oIntersecPos.Set x, y, z
            Set oVec = oPos.Subtract(oIntersecPos)
            
            Dim curDepth As Double, deltaDepth As Double
            'get the current depth along cope y vector
            curDepth = oVec.Dot(oCopeYvec)
                        
            dblCopedepth = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("Depth").Value
            
            'curdepth is expected to be less than dblcopedepth
            deltaDepth = dblCopedepth - Abs(curDepth)
            If Abs(deltaDepth) > distTol Then
                Set oVec = oCopeYvec.Clone()
                If curDepth < 0 Then ' we are computing top cope
                    oVec.length = -deltaDepth
                Else
                    oVec.length = deltaDepth
                End If
    
                Set oPos = oPos.Offset(oVec)
            End If
            'create line through oPos  parallel to supped axis
            Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, oPos.x, oPos.y, oPos.z, oCopeXvec.x, oCopeXvec.y, oCopeXvec.z, 1#)
            oLine.Infinite = True
            Set oSurf = oCutbackPlane
            oSurf.Intersect oLine, colElms, code
            If Not colElms Is Nothing Then
                If colElms.count > 0 Then
                    Set oIntrSecPoint = colElms.Item(1)
                    oIntrSecPoint.GetPoint x, y, z
                    oCopeWebPos.Set x, y, z
                    Set oVec = oPos.Subtract(oCopeWebPos)
                    dblCopeLength = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("Length").Value
                    
                    oVec.length = dblCopeLength
                    Set oPos = oCopeWebPos.Offset(oVec)
                    ' modify cope plane root point to adjust for cope length and cope depth changes
                    oCopePlane.SetRootPoint oPos.x, oPos.y, oPos.z
                    
                    If bIsSuppedAxisInPositiveZ Then
                        Set oTopCopePos = oPos.Clone()
                    Else
                        Set oBottCopePos = oPos.Clone()
                    End If
                    Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, oCopeWebPos.x, oCopeWebPos.y, oCopeWebPos.z, oCopeYvec.x, oCopeYvec.y, oCopeYvec.z, 1#)
                    'make the line infinite
                    oLine.Infinite = True
                    If bIsSuppedAxisInPositiveZ Then
                        oRightTopOfBottomFlange.Get x, y, z
                    Else
                        oRightBottomOfTopFlange.Get x, y, z
                    End If
                    Set oSurf = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, x, y, z, oSuppingZvec.x, oSuppingZvec.y, oSuppingZvec.z)
                    oSurf.Intersect oLine, colElms, code
                    If Not colElms Is Nothing Then
                        If colElms.count > 0 Then
                            Set oIntrSecPoint = colElms.Item(1)
                            oIntrSecPoint.GetPoint x, y, z
                            Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, x, y, z, oCopeXvec.x, oCopeXvec.y, oCopeXvec.z, 1#)
                            oLine.Infinite = True
                            Set oSurf = oCopePlane
                            oSurf.Intersect oLine, colElms, code
                            If Not colElms Is Nothing Then
                                If colElms.count > 0 Then
                                    Set oIntrSecPoint = colElms.Item(1)
                                    oIntrSecPoint.GetPoint x, y, z
                                    oPos.Set x, y, z
                                    GetCopePoint oSupping, oSupped, iEnd, oCopePlane, Not bIsSuppedAxisInPositiveZ, oIntrSecPoint
                                    oIntrSecPoint.GetPoint x, y, z
                                    oIntersecPos.Set x, y, z
                                   
                                    Set oVec = oPos.Subtract(oIntersecPos)
                                    'get the current depth along cope y vector
                                    curDepth = oVec.Dot(oCopeYvec)
                                   
                                    deltaDepth = dblCopedepth - curDepth
                                    If Abs(deltaDepth) > distTol Then
                                        Set oVec = oCopeYvec.Clone()
                                        If curDepth < 0 Then ' we are computing top cope
                                            oVec.length = -deltaDepth
                                        Else
                                            oVec.length = deltaDepth
                                        End If
                                        Set oPos = oPos.Offset(oVec)
                                    End If

                                    If bIsSuppedAxisInPositiveZ Then
                                        Set oBottCopePos = oPos
                                    Else
                                        Set oTopCopePos = oPos
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
  For i = 1 To 4
    Set oCopePos(i) = New DPosition
  Next i
  If bTop Then
    Set oCopePos(1) = oTopCopePos
    intFlag = 1
  Else
    Set oCopePos(1) = oBottCopePos
    intFlag = -1
  End If
  
  'let us convert oCopPos(1) to local coordinates
  oMat.LoadIdentity
  
  oMat.IndexValue(0) = oCopeXvec.x
  oMat.IndexValue(1) = oCopeXvec.y
  oMat.IndexValue(2) = oCopeXvec.z

  oMat.IndexValue(4) = oCopeYvec.x
  oMat.IndexValue(5) = oCopeYvec.y
  oMat.IndexValue(6) = oCopeYvec.z

  oMat.IndexValue(8) = oCopeZvec.x
  oMat.IndexValue(9) = oCopeZvec.y
  oMat.IndexValue(10) = oCopeZvec.z
  
  'make oCopepos(1) the origin of the coord system
  oMat.IndexValue(12) = oCopePos(1).x
  oMat.IndexValue(13) = oCopePos(1).y
  oMat.IndexValue(14) = oCopePos(1).z
  
  'convert to local
  oMat.Invert
  
  Set oCopePos(1) = oMat.TransformPosition(oCopePos(1))
  
  oCopePos(1).z = 0 'not needed as position already at the origin. But set to zero as there could be a tolerance issue
  

  oCopePos(2).x = 0
  oCopePos(2).y = intFlag * LARGE_EDGE
  oCopePos(2).z = 0
  
  oCopePos(3).x = -LARGE_EDGE
  oCopePos(3).y = oCopePos(2).y
  oCopePos(3).z = 0

  oCopePos(4).x = -LARGE_EDGE
  oCopePos(4).y = oCopePos(1).y
  oCopePos(4).z = 0
  Dim curveElms As IJElements
  Dim oLine3d(1 To 4) As Line3d
  Dim oCmplx As ComplexString3d
  Dim oIJGeometryMisc As IJGeometryMisc
  Dim oProfile2dhelper As IJDProfile2dHelper
  Dim oArrayOfids(0 To 3) As Long
  
  Set curveElms = New JObjectCollection
  For i = 1 To 4
    Dim nextIdx As Integer
    nextIdx = i Mod 4
    Set oLine3d(i) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oCopePos(i).x, oCopePos(i).y, oCopePos(i).z, _
    oCopePos(nextIdx + 1).x, oCopePos(nextIdx + 1).y, oCopePos(nextIdx + 1).z)
    curveElms.Add oLine3d(i)
    Set oLine3d(i) = Nothing
  Next i

  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
  
  Dim oAxisLine As IJLine
  Dim lengthSupped#
  
  lengthSupped = oSupped.Axis.length
  'we need create an axis curve of  length (enough to cut entire width of supported member)
  'for creating the solid cutter at the same time not cutting the other end of the supped member
  'so a length equal to that of the supped member is good
  
  'the start will be some distance,say lengthSupped/2, away from the origin of the coord system
  'the axis will be in the Z direction
  oMat.Invert
  x = oMat.IndexValue(12) - lengthSupped * 0.5 * oMat.IndexValue(8)
  y = oMat.IndexValue(13) - lengthSupped * 0.5 * oMat.IndexValue(9)
  z = oMat.IndexValue(14) - lengthSupped * 0.5 * oMat.IndexValue(10)
  

  Set oAxisLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, x, y, z, oMat.IndexValue(8), oMat.IndexValue(9), _
                oMat.IndexValue(10), lengthSupped)

  'create solid cutter by projecting the cutting shape along the axis
  CreateSolidCutter oRescMgr, oSupping, oSupped, iEnd, oCmplx, oCutter, oAxisLine, oCopeYvec
  
  For i = 1 To 4
    Set oCopePos(i) = Nothing
  Next i
  Exit Sub
ErrorHandler:    HandleError MODULE, MT
End Sub

'Public Sub CreateCornerCope(ByVal CopeType As Integer, oSupped As ISPSMemberPartPrismatic, oSuppingPlane1 As IJPlane, oSuppingPlane2 As IJPlane, oAttribsCAO As IJDAttributes, oRescMgr As IUnknown, ByRef oWireBody As IJWireBody)
'Const MT = "CreateCornerCope"
'  On Error GoTo ErrorHandler
'
'  Dim CopeLength As Double, CopeDepth As Double, SideClr As Double, InsideClr As Double, Radius As Double, RoundIncrement As Double, RadiusType As Long
'  Dim IntLine As Line3d
'  Dim pGeometryFactory As New GeometryFactory
'  Dim oCmplx As ComplexString3d
'  Dim oIJGeometryMisc As IJGeometryMisc
'  Dim oProfile2dhelper As IJDProfile2dHelper
'  Dim oArrayOfids(12) As Long
'  Dim pt1x As Double, pt1y As Double, pt1z As Double, pt2x As Double, pt2y As Double, pt2z As Double
'  Dim n1x As Double, n1y As Double, n1z As Double, n2x As Double, n2y As Double, n2z As Double
'  Dim oSurf1 As IJSurface, oSurf2 As IJSurface
'  Dim code As Geom3dIntersectConstants
'  Dim intElms As IJElements
'  Dim pt1 As IJDPosition, pt1p As IJDPosition, pt2 As IJDPosition
'  Dim Shp1Vec As IJDVector, Shp2Vec As IJDVector, Clr1Vec As IJDVector, Clr2Vec As IJDVector, planevec1 As IJDVector, planevec2 As IJDVector
'  Dim InfLine(0 To 1) As IJDPosition, elmCount As Integer
'  Dim SquareEnd As Boolean, PlanesConcave As Boolean, idx As Integer
'  Dim curveElms As IJElements
'  Dim WhichEnd As SPSMemberAxisPortIndex, CopeSide As Integer, CopeEdge As Integer
'  Dim WebRadius As Double, WebRadiusType As Long
'
'  Set Shp1Vec = New DVector
'  Set Shp2Vec = New DVector
'  Set Clr1Vec = New DVector
'  Set Clr2Vec = New DVector
'  Set planevec1 = New DVector
'  Set planevec2 = New DVector
'  Set pt1 = New DPosition
'  Set pt2 = New DPosition
'  Set pt1p = New DPosition
'  Set InfLine(0) = New DPosition
'  Set InfLine(1) = New DPosition
'
'  SideClr = oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("CopeLengthClearance").Value
'  InsideClr = oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("CopeDepthClearance").Value
'  Radius = oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("CopeRadius").Value
'  RadiusType = oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("CopeRadiusType").Value
'  RoundIncrement = oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("RoundingIncrement").Value
'  WebRadius = oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("WebRadius").Value
'  WebRadiusType = oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("WebRadiusType").Value
'  SquareEnd = 1 ' currently always set to TRUE
'
'  ' now get a point on the line of intersection of the two planes
'  oSuppingPlane1.GetRootPoint pt1x, pt1y, pt1z
'  pt1.Set pt1x, pt1y, pt1z
'  oSuppingPlane2.GetRootPoint pt2x, pt2y, pt2z
'  pt2.Set pt2x, pt2y, pt2z
'  oSuppingPlane1.GetNormal n1x, n1y, n1z
'  planevec1.Set n1x, n1y, n1z
'  oSuppingPlane2.GetNormal n2x, n2y, n2z
'  planevec2.Set n2x, n2y, n2z
'  Set oSurf1 = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, pt1x, pt1y, pt1z, n1x, n1y, n1z)
'  Set oSurf2 = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, pt2x, pt2y, pt2z, n2x, n2y, n2z)
'  oSurf1.Intersect oSurf2, intElms, code
'  If intElms Is Nothing Then
'    GoTo ErrorHandler
'  ElseIf ((Not intElms Is Nothing) And intElms.Count <= 0) Then
'    GoTo ErrorHandler
'  End If
'  Set IntLine = intElms.Item(1)
'  IntLine.GetStartPoint pt1x, pt1y, pt1z
'  InfLine(0).Set pt1x, pt1y, pt1z
'  IntLine.GetEndPoint pt1x, pt1y, pt1z
'  InfLine(1).Set pt1x, pt1y, pt1z
'  ' make finite line
'  MakeLineFinite InfLine
'  ProjectPtOnLine pt1, InfLine, pt1p
'
'  ' now get the vectors going in and coming out of target member, for each plane
'  GetShapeAndClearanceDirections oSupped, oSuppingPlane1, oSuppingPlane2, CopeType, PlanesConcave, Shp1Vec, Clr1Vec, Shp2Vec, Clr2Vec
'  Shp1Vec.length = 1#
'  Shp2Vec.length = 1#
'  Clr1Vec.length = 1#
'  Clr2Vec.length = 1#
'  CreateCopeCuttingShapeAndComputeOutputs oSupped, pt1p, Shp1Vec, Shp2Vec, SideClr, Clr1Vec, InsideClr, Clr2Vec, RoundIncrement, SquareEnd, PlanesConcave, Radius, RadiusType, curveElms, CopeLength, CopeDepth, CopeType, CopeSide, CopeEdge, WhichEnd
'  elmCount = curveElms.Count
'  Set oCmplx = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
'  Set curveElms = Nothing
'  Set oIJGeometryMisc = New DGeomOpsMisc
'  If oWireBody Is Nothing Then
'    oIJGeometryMisc.CreateModelGeometryFromGType oRescMgr, oCmplx, Nothing, oWireBody
'    If oWireBody Is Nothing Then GoTo ErrorHandler
'  Else
'    oIJGeometryMisc.ModifyModelGeometryFromGType oCmplx, oWireBody
'  End If
'  For idx = 0 To elmCount - 1
'    oArrayOfids(idx) = idx + 1
'  Next idx
'
'  Set oProfile2dhelper = New DProfile2dHelper
'  oProfile2dhelper.SetIntegerAttributesOnWireBodyEdges oWireBody, "JSXid", elmCount, oArrayOfids(0)
'
'  ' now set all the outputs for the web corner cope
'  oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("CopeLength").Value = CopeLength
'  oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("CopeDepth").Value = CopeDepth
'  oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("CopeEdge").Value = CopeEdge
'  oAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope").Item("CopeEnd").Value = WhichEnd
'  Exit Sub
'ErrorHandler:    HandleError MODULE, MT
'End Sub
'Public Sub GetShapeAndClearanceDirections(ByVal oSuppedPart As ISPSMemberPartPrismatic, oSuppingPlane1 As IJPlane, oSuppingPlane2 As IJPlane, ByVal CopeType As Integer, ByRef PlanesConcave As Boolean, ByRef Shp1Vec As IJDVector, ByRef Clr1Vec As IJDVector, ByRef Shp2Vec As IJDVector, ByRef Clr2Vec As IJDVector)
'Const MT = "GetShapeAndClearanceDirections"
'    On Error GoTo ErrorHandler
'    Dim oCurve1 As IJCurve, oCurve2 As IJCurve
'    Dim pt1 As IJDPosition, pt11 As IJDPosition, pt2 As IJDPosition, pt22 As IJDPosition, IntPt As IJDPosition
'    Dim oGeomFact As New GeometryFactory
'    Dim code As Geom3dIntersectConstants
'    Dim NumIntersects As Long, overlap As Long, status As Integer
'    Dim IntPoints() As Double
'    Dim oLine As IJLine, px As Double, py As Double, pz As Double
'    Dim planevec1 As IJDVector, planevec2 As IJDVector, intElms As IJElements, TmpVec As IJDVector, TmpVec1 As IJDVector, TmpVec2 As IJDVector
'    Dim CutbackEnd As SPSMemberAxisPortIndex, CutPlane As Integer
'    Dim oSurf1 As IJSurface, oSurf2 As IJSurface
'    Dim InfLine(0 To 1) As IJDPosition, OtherEnd As IJDPosition, InterSecPoint As IJPoint
'    Dim IntLine As Line3d, TmpPt1 As IJDPosition, TmpPt2 As IJDPosition
'    Dim ShpElms As IJElements
'    Dim LenVec As IJDVector, DepVec As IJDVector, oStartPos(0 To 3) As IJDPosition, idx As Integer
'
'    Set pt1 = New DPosition
'    Set pt11 = New DPosition
'    Set pt2 = New DPosition
'    Set pt22 = New DPosition
'    Set planevec1 = New DVector
'    Set planevec2 = New DVector
'    Set TmpVec = New DVector
'    Set IntPt = New DPosition
'    Set OtherEnd = New DPosition
'    Set InfLine(0) = New DPosition
'    Set InfLine(1) = New DPosition
'    Set TmpVec1 = New DVector
'    Set TmpVec2 = New DVector
'    Set TmpPt1 = New DPosition
'    Set TmpPt2 = New DPosition
'    Set LenVec = New DVector
'    Set DepVec = New DVector
'    Set oStartPos(0) = New DPosition
'    Set oStartPos(1) = New DPosition
'    Set oStartPos(2) = New DPosition
'    Set oStartPos(3) = New DPosition
'
'    PlanesConcave = False
'    oSuppingPlane1.GetBoundaries ShpElms
'    status = ComputeCentroidOfShape(ShpElms, px, py, pz)
'    If status = 0 Then ' success
'    Else
'        oSuppingPlane1.GetRootPoint px, py, pz
'    End If
'    Set ShpElms = Nothing
'    pt1.Set px, py, pz
'    oSuppingPlane1.GetNormal px, py, pz
'    planevec1.Set px, py, pz
'    oSuppingPlane2.GetBoundaries ShpElms
'    status = ComputeCentroidOfShape(ShpElms, px, py, pz)
'    If status = 0 Then ' success
'    Else
'        oSuppingPlane2.GetRootPoint px, py, pz
'    End If
'    Set ShpElms = Nothing
'    pt2.Set px, py, pz
'    oSuppingPlane2.GetNormal px, py, pz
'    planevec2.Set px, py, pz
'
'    Set oSurf1 = oGeomFact.Planes3d.CreateByPointNormal(Nothing, pt1.x, pt1.y, pt1.z, planevec1.x, planevec1.y, planevec1.z)
'    Set oSurf2 = oGeomFact.Planes3d.CreateByPointNormal(Nothing, pt2.x, pt2.y, pt2.z, planevec2.x, planevec2.y, planevec2.z)
'    oSurf1.Intersect oSurf2, intElms, code
'    If intElms Is Nothing Then
'        GoTo ErrorHandler
'    ElseIf ((Not intElms Is Nothing) And intElms.Count <= 0) Then
'        GoTo ErrorHandler
'    End If
'    Set oSurf1 = Nothing
'    Set oSurf2 = Nothing
'    Set IntLine = intElms.Item(1)
'    IntLine.GetStartPoint px, py, pz
'    InfLine(0).Set px, py, pz
'    IntLine.GetEndPoint px, py, pz
'    InfLine(1).Set px, py, pz
'    ' make finite line
'    MakeLineFinite InfLine
'
'    If IsSurfaceTrimCutbackNeeded(oSuppedPart, oSuppingPlane1) Then
'        CutbackEnd = GetCutbackEnd(oSuppedPart, oSuppingPlane1)
'        CutPlane = 1
'    ElseIf IsSurfaceTrimCutbackNeeded(oSuppedPart, oSuppingPlane2) Then
'        CutbackEnd = GetCutbackEnd(oSuppedPart, oSuppingPlane2)
'        CutPlane = 2
'    Else
'        GoTo ErrorHandler
'    End If
'    Set oLine = oSuppedPart
'    If CutbackEnd = SPSMemberAxisStart Then
'        oLine.GetEndPoint px, py, pz
'    Else
'        oLine.GetStartPoint px, py, pz
'    End If
'    pt11.Set px, py, pz
'    Set oLine = Nothing
'    OtherEnd.Set pt11.x, pt11.y, pt11.z
'    If CutPlane = 1 Then
'        Set pt22 = ProjectPosToPlane(pt11, pt1, planevec1)
'        Clr1Vec.Set pt11.x - pt22.x, pt11.y - pt22.y, pt11.z - pt22.z
'        Clr1Vec.length = 1#
'
'        Set TmpPt1 = ProjectPosToPlane(pt1, pt2, planevec2)
'        If TmpPt1.DistPt(pt1) <= 0.0001 Then
'            Set TmpPt1 = ProjectPosToPlane(pt22, pt2, planevec2)
'            If TmpPt1.DistPt(pt22) <= 0.0001 Then
'                pt22.Set pt1.x + planevec2.x, pt1.y + planevec2.y, pt1.z + planevec2.z
'                pt11.Set pt22.x + planevec1.x, pt22.y + planevec1.y, pt22.z + planevec1.z
'            End If
'        Else
'            pt22.Set pt1.x, pt1.y, pt1.z
'            pt11.Set pt22.x + planevec1.x, pt22.y + planevec1.y, pt22.z + planevec1.z
'        End If
'
'        ProjectPtOnLine pt22, InfLine, TmpPt1
'        ProjectPtOnLine pt2, InfLine, TmpPt2
'        TmpVec.Set TmpPt1.x - TmpPt2.x, TmpPt1.y - TmpPt2.y, TmpPt1.z - TmpPt2.z
'        TmpPt2.Set pt2.x + TmpVec.x, pt2.y + TmpVec.y, pt2.z + TmpVec.z
'
'        Set oLine = oGeomFact.Lines3d.CreateBy2Points(Nothing, pt22.x, pt22.y, pt22.z, pt11.x, pt11.y, pt11.z)
'        oLine.Infinite = True
'        Set oCurve1 = oLine
'        Set oLine = Nothing
'        Set oLine = oGeomFact.Lines3d.CreateByPtVectLength(Nothing, TmpPt2.x, TmpPt2.y, TmpPt2.z, planevec2.x, planevec2.y, planevec2.z, LARGE_EDGE)
'        oLine.Infinite = True
'        Set oCurve2 = oLine
'        Set oLine = Nothing
'        oCurve1.Intersect oCurve2, NumIntersects, IntPoints, overlap, ISECT_UNKNOWN
'        If NumIntersects = 1 Then
'            pt11.Set IntPoints(0), IntPoints(1), IntPoints(2)
'            TmpVec.Set pt11.x - pt22.x, pt11.y - pt22.y, pt11.z - pt22.z
'            TmpVec.length = 1#
'            ' now project int point on plane 2
'            Set pt22 = ProjectPosToPlane(pt11, pt2, planevec2)
'            If Abs(TmpVec.x - Clr1Vec.x) <= 0.0001 And Abs(TmpVec.y - Clr1Vec.y) <= 0.0001 And Abs(TmpVec.z - Clr1Vec.z) <= 0.0001 Then
'                Clr2Vec.Set pt11.x - pt22.x, pt11.y - pt22.y, pt11.z - pt22.z
'            Else ' int point is on opposite side of clr1vec
'                Clr2Vec.Set pt22.x - pt11.x, pt22.y - pt11.y, pt22.z - pt11.z
'            End If
'            Clr2Vec.length = 1#
'        Else
'            GoTo ErrorHandler
'        End If
'    Else
'        Set pt22 = ProjectPosToPlane(pt11, pt2, planevec2)
'        Clr2Vec.Set pt11.x - pt22.x, pt11.y - pt22.y, pt11.z - pt22.z
'        Clr2Vec.length = 1#
'
'        Set TmpPt1 = ProjectPosToPlane(pt2, pt1, planevec1)
'        If TmpPt1.DistPt(pt2) <= 0.0001 Then
'            Set TmpPt1 = ProjectPosToPlane(pt22, pt1, planevec1)
'            If TmpPt1.DistPt(pt22) <= 0.0001 Then
'                pt22.Set pt2.x + planevec1.x, pt2.y + planevec1.y, pt2.z + planevec1.z
'                pt11.Set pt22.x + planevec2.x, pt22.y + planevec2.y, pt22.z + planevec2.z
'            End If
'        Else
'            pt22.Set pt2.x, pt2.y, pt2.z
'            pt11.Set pt22.x + planevec2.x, pt22.y + planevec2.y, pt22.z + planevec2.z
'        End If
'
'        ProjectPtOnLine pt22, InfLine, TmpPt1
'        ProjectPtOnLine pt1, InfLine, TmpPt2
'        TmpVec.Set TmpPt1.x - TmpPt2.x, TmpPt1.y - TmpPt2.y, TmpPt1.z - TmpPt2.z
'        TmpPt2.Set pt1.x + TmpVec.x, pt1.y + TmpVec.y, pt1.z + TmpVec.z
'        Set oLine = oGeomFact.Lines3d.CreateBy2Points(Nothing, pt22.x, pt22.y, pt22.z, pt11.x, pt11.y, pt11.z)
'        oLine.Infinite = True
'        Set oCurve1 = oLine
'        Set oLine = Nothing
'        Set oLine = oGeomFact.Lines3d.CreateByPtVectLength(Nothing, TmpPt2.x, TmpPt2.y, TmpPt2.z, planevec1.x, planevec1.y, planevec1.z, LARGE_EDGE)
'        oLine.Infinite = True
'        Set oCurve2 = oLine
'        Set oLine = Nothing
'        oCurve1.Intersect oCurve2, NumIntersects, IntPoints, overlap, ISECT_UNKNOWN
'        If NumIntersects = 1 Then
'            pt11.Set IntPoints(0), IntPoints(1), IntPoints(2)
'            TmpVec.Set pt11.x - pt22.x, pt11.y - pt22.y, pt11.z - pt22.z
'            TmpVec.length = 1#
'            ' now project int point on plane 1
'            Set pt22 = ProjectPosToPlane(pt11, pt1, planevec1)
'            If Abs(TmpVec.x - Clr2Vec.x) <= 0.0001 And Abs(TmpVec.y - Clr2Vec.y) <= 0.0001 And Abs(TmpVec.z - Clr2Vec.z) <= 0.0001 Then
'                Clr1Vec.Set pt11.x - pt22.x, pt11.y - pt22.y, pt11.z - pt22.z
'            Else ' int point is on opposite side of clr2vec
'                Clr1Vec.Set pt22.x - pt11.x, pt22.y - pt11.y, pt22.z - pt11.z
'            End If
'            Clr1Vec.length = 1#
'        Else
'            GoTo ErrorHandler
'        End If
'    End If
'
'    ' now get the shape1 and shape2 vectors
'    IntPt.Set pt11.x, pt11.y, pt11.z
'    ProjectPtOnLine IntPt, InfLine, pt11
'    TmpVec.Set InfLine(0).x - InfLine(1).x, InfLine(0).y - InfLine(1).y, InfLine(0).z - InfLine(1).z
'    TmpVec.length = 1#
'    Set TmpVec1 = TmpVec.Cross(Clr1Vec)
'    TmpVec1.length = 1#
'    Set TmpVec2 = TmpVec.Cross(Clr2Vec)
'    TmpVec2.length = 1#
'    InfLine(0).Set pt11.x, pt11.y, pt11.z
'    InfLine(1).Set pt11.x + LARGE_EDGE * TmpVec1.x, pt11.y + LARGE_EDGE * TmpVec1.y, pt11.z + LARGE_EDGE * TmpVec1.z
'    ProjectPtOnLine IntPt, InfLine, pt22
'    Shp1Vec.Set pt22.x - pt11.x, pt22.y - pt11.y, pt22.z - pt11.z
'    Shp1Vec.length = 1#
'    InfLine(1).Set pt11.x + LARGE_EDGE * TmpVec2.x, pt11.y + LARGE_EDGE * TmpVec2.y, pt11.z + LARGE_EDGE * TmpVec2.z
'    ProjectPtOnLine IntPt, InfLine, pt22
'    Shp2Vec.Set pt22.x - pt11.x, pt22.y - pt11.y, pt22.z - pt11.z
'    Shp2Vec.length = 1#
'
'    ' now determine whether the planes are concave or convex
'    InfLine(0).Set IntPt.x, IntPt.y, IntPt.z
'    If CutPlane = 1 Then
'        Set oSurf1 = oGeomFact.Planes3d.CreateByPointNormal(Nothing, pt1.x, pt1.y, pt1.z, planevec1.x, planevec1.y, planevec1.z)
'        InfLine(1).Set IntPt.x + LARGE_EDGE * Clr1Vec.x, IntPt.y + LARGE_EDGE * Clr1Vec.y, IntPt.z + LARGE_EDGE * Clr1Vec.z
'    Else
'        Set oSurf1 = oGeomFact.Planes3d.CreateByPointNormal(Nothing, pt2.x, pt2.y, pt2.z, planevec2.x, planevec2.y, planevec2.z)
'        InfLine(1).Set IntPt.x + LARGE_EDGE * Clr2Vec.x, IntPt.y + LARGE_EDGE * Clr2Vec.y, IntPt.z + LARGE_EDGE * Clr2Vec.z
'    End If
'    Set oLine = oGeomFact.Lines3d.CreateBy2Points(Nothing, InfLine(0).x, InfLine(0).y, InfLine(0).z, InfLine(1).x, InfLine(1).y, InfLine(1).z)
'    oLine.Infinite = True
'    oSurf1.Intersect oLine, intElms, code
'    If intElms Is Nothing Then
'        GoTo ErrorHandler
'    ElseIf ((Not intElms Is Nothing) And intElms.Count <= 0) Then
'        GoTo ErrorHandler
'    End If
'    Set InterSecPoint = intElms.Item(1)
'    InterSecPoint.GetPoint px, py, pz
'    pt11.Set px, py, pz
'    ProjectPtOnLine OtherEnd, InfLine, TmpPt2
'    TmpVec1.Set TmpPt2.x - pt11.x, TmpPt2.y - pt11.y, TmpPt2.z - pt11.z
'    TmpVec1.length = 1#
'    TmpVec2.Set IntPt.x - pt11.x, IntPt.y - pt11.y, IntPt.z - pt11.z
'    TmpVec2.length = 1#
'    If Abs(TmpVec2.x - TmpVec1.x) <= 0.0001 And Abs(TmpVec2.y - TmpVec1.y) <= 0.0001 And Abs(TmpVec2.z - TmpVec1.z) <= 0.0001 Then ' both are same
'        PlanesConcave = True
'    Else ' intpt and other end are on either sides of the planes
'        PlanesConcave = False
'    End If
'
'    ' now make the clearance directions parallel/perpendicular to member axis
'    Set oLine = oSuppedPart
'    oLine.GetStartPoint px, py, pz
'    TmpPt1.Set px, py, pz
'    oLine.GetEndPoint px, py, pz
'    TmpPt2.Set px, py, pz
'    LenVec.Set TmpPt2.x - TmpPt1.x, TmpPt2.y - TmpPt1.y, TmpPt2.z - TmpPt1.z
'    LenVec.length = 1#
'    GetBoundingRectangleForMemberEnd oSuppedPart, SPSMemberAxisStart, oStartPos
'    If CopeType = 1 Then ' webcope
'        DepVec.Set oStartPos(3).x - oStartPos(0).x, oStartPos(3).y - oStartPos(0).y, oStartPos(3).z - oStartPos(0).z
'    Else ' flange cope
'        DepVec.Set oStartPos(1).x - oStartPos(0).x, oStartPos(1).y - oStartPos(0).y, oStartPos(1).z - oStartPos(0).z
'    End If
'    DepVec.length = 1#
'    InfLine(0).Set OtherEnd.x, OtherEnd.y, OtherEnd.z
'    InfLine(1).Set OtherEnd.x + LARGE_EDGE * LenVec.x, OtherEnd.y + LARGE_EDGE * LenVec.y, OtherEnd.z + LARGE_EDGE * LenVec.z
'    TmpPt1.Set OtherEnd.x + LARGE_EDGE * Clr1Vec.x, OtherEnd.y + LARGE_EDGE * Clr1Vec.y, OtherEnd.z + LARGE_EDGE * Clr1Vec.z
'    ProjectPtOnLine TmpPt1, InfLine, TmpPt2
'    Clr1Vec.Set TmpPt2.x - OtherEnd.x, TmpPt2.y - OtherEnd.y, TmpPt2.z - OtherEnd.z
'    Clr1Vec.length = 1#
'
'    InfLine(1).Set OtherEnd.x + LARGE_EDGE * DepVec.x, OtherEnd.y + LARGE_EDGE * DepVec.y, OtherEnd.z + LARGE_EDGE * DepVec.z
'    TmpPt1.Set OtherEnd.x + LARGE_EDGE * Clr2Vec.x, OtherEnd.y + LARGE_EDGE * Clr2Vec.y, OtherEnd.z + LARGE_EDGE * Clr2Vec.z
'    ProjectPtOnLine TmpPt1, InfLine, TmpPt2
'    Clr2Vec.Set TmpPt2.x - OtherEnd.x, TmpPt2.y - OtherEnd.y, TmpPt2.z - OtherEnd.z
'    Clr2Vec.length = 1#
'    For idx = 0 To 3
'        oStartPos(idx) = Nothing
'    Next idx
'    Set oLine = Nothing
'
'    Set TmpPt1 = Nothing
'    Set TmpPt2 = Nothing
'    Set oLine = Nothing
'    Set oSurf1 = Nothing
'    Set pt11 = Nothing
'    Set pt22 = Nothing
'    Set pt1 = Nothing
'    Set pt2 = Nothing
'    Set planevec1 = Nothing
'    Set planevec2 = Nothing
'    Set TmpVec = Nothing
'    Set InfLine(0) = Nothing
'    Set InfLine(1) = Nothing
'    Set TmpVec1 = Nothing
'    Set TmpVec2 = Nothing
'    Set OtherEnd = Nothing
'    Set oGeomFact = Nothing
'    Set oCurve1 = Nothing
'    Set oCurve2 = Nothing
'    Set IntPt = Nothing
'    Set LenVec = Nothing
'    Set DepVec = Nothing
'    Exit Sub
'
'ErrorHandler:
'    HandleError MODULE, MT
'End Sub

Public Sub ProjectPtOnLine(ByVal pt1 As IJDPosition, ByRef Line() As IJDPosition, ByRef pt1proj As IJDPosition)
    Dim vec1 As IJDVector, vec2 As IJDVector
    Dim s As Double, dot_prod As Double
    Dim px As Double, py As Double, pz As Double
    
    Set vec1 = New DVector
    Set vec2 = New DVector
    
    Set vec1 = Line(1).Subtract(Line(0))
    dot_prod = vec1.Dot(vec1)
    Set vec2 = pt1.Subtract(Line(0))
    s = 0#
    If dot_prod <> 0# Then
       s = vec2.Dot(vec1) / dot_prod
    End If
    px = Line(0).x + (s * vec1.x)
    py = Line(0).y + (s * vec1.y)
    pz = Line(0).z + (s * vec1.z)
    pt1proj.Set px, py, pz
    
    Set vec1 = Nothing
    Set vec2 = Nothing
End Sub
'Public Function IsCornerCopeNeeded(oSuppedPart As ISPSMemberPartPrismatic, oSuppingPlane1 As IJPlane, oSuppingPlane2 As IJPlane) As Boolean
'    Const METHOD = "IsCornerCopeNeeded"
'    On Error GoTo ErrorHandler
'    Dim bIsNeeded As Boolean
'    Dim oSurf1 As IJSurface, oSurf2 As IJSurface, intElms As IJElements
'    Dim code As Geom3dIntersectConstants, IntLine As Line3d, i As Integer, nextIdx As Integer
'    Dim oStartPos(0 To 3) As IJDPosition, oEndPos(0 To 3) As IJDPosition
'    Dim pGeometryFactory As New GeometryFactory
'    Dim PlanePoints(12) As Double
'    Dim px As Double, py As Double, pz As Double, nX As Double, nY As Double, nZ As Double
'
'    For i = 0 To 3
'        Set oStartPos(i) = New DPosition
'        Set oEndPos(i) = New DPosition
'    Next i
'    oSuppingPlane1.GetRootPoint px, py, pz
'    oSuppingPlane1.GetNormal nX, nY, nZ
'    Set oSurf1 = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, px, py, pz, nX, nY, nZ)
'    oSuppingPlane2.GetRootPoint px, py, pz
'    oSuppingPlane2.GetNormal nX, nY, nZ
'    Set oSurf2 = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, px, py, pz, nX, nY, nZ)
'    oSurf1.Intersect oSurf2, intElms, code
'    If intElms Is Nothing Then
'        GoTo ErrorHandler
'    ElseIf ((Not intElms Is Nothing) And intElms.Count <= 0) Then
'        GoTo ErrorHandler
'    End If
'    Set IntLine = intElms.Item(1)
'    IntLine.Infinite = True
'    Set oSurf1 = Nothing
'    Set oSurf2 = Nothing
'
'    ' now check intersection of the line with the supped part
'    bIsNeeded = False
'    GetBoundingRectangleForMemberEnd oSuppedPart, SPSMemberAxisStart, oStartPos
'    GetBoundingRectangleForMemberEnd oSuppedPart, SPSMemberAxisEnd, oEndPos
'    ' check for intersection with start endshape
'    For i = 0 To 3
'        PlanePoints(i * 3) = oStartPos(i).x
'        PlanePoints(i * 3 + 1) = oStartPos(i).y
'        PlanePoints(i * 3 + 2) = oStartPos(i).z
'    Next i
'    Set oSurf1 = pGeometryFactory.Planes3d.CreateByPoints(Nothing, 4, PlanePoints)
'    oSurf1.Intersect IntLine, intElms, code
'    If Not intElms Is Nothing Then
'        If intElms.Count > 0 Then
'            bIsNeeded = True
'        End If
'    End If
'    If bIsNeeded = False Then
'        ' now check with end endshape
'        For i = 0 To 3
'            PlanePoints(i * 3) = oEndPos(i).x
'            PlanePoints(i * 3 + 1) = oEndPos(i).y
'            PlanePoints(i * 3 + 2) = oEndPos(i).z
'        Next i
'        Set oSurf1 = pGeometryFactory.Planes3d.CreateByPoints(Nothing, 4, PlanePoints)
'        oSurf1.Intersect IntLine, intElms, code
'        If Not intElms Is Nothing Then
'            If intElms.Count > 0 Then
'                bIsNeeded = True
'            End If
'        End If
'    End If
'    If bIsNeeded = False Then
'        ' now check with side faces
'        For i = 0 To 3
'            nextIdx = (i + 1) Mod 4
'            PlanePoints(0) = oStartPos(i).x
'            PlanePoints(1) = oStartPos(i).y
'            PlanePoints(2) = oStartPos(i).z
'            PlanePoints(3) = oStartPos(nextIdx).x
'            PlanePoints(4) = oStartPos(nextIdx).y
'            PlanePoints(5) = oStartPos(nextIdx).z
'            PlanePoints(6) = oEndPos(nextIdx).x
'            PlanePoints(7) = oEndPos(nextIdx).y
'            PlanePoints(8) = oEndPos(nextIdx).z
'            PlanePoints(9) = oEndPos(i).x
'            PlanePoints(10) = oEndPos(i).y
'            PlanePoints(11) = oEndPos(i).z
'            Set oSurf1 = pGeometryFactory.Planes3d.CreateByPoints(Nothing, 4, PlanePoints)
'            oSurf1.Intersect IntLine, intElms, code
'            If Not intElms Is Nothing Then
'                If intElms.Count > 0 Then
'                    bIsNeeded = True
'                    Exit For
'                End If
'            End If
'        Next i
'    End If
'
'    For i = 0 To 3
'        Set oStartPos(i) = Nothing
'        Set oEndPos(i) = Nothing
'    Next i
'    IsCornerCopeNeeded = bIsNeeded
'    Exit Function
'
'ErrorHandler:
'    IsCornerCopeNeeded = False
'    HandleError MODULE, METHOD
'End Function
Public Function GetCutbackEnd(oPart As ISPSMemberPartPrismatic, oSuppingSurf As IJSurface) As SPSMemberAxisPortIndex
    Const METHOD = "GetCutbackEnd"
    On Error GoTo ErrorHandler

    Dim oSurf1 As IJSurface, oLine As IJLine, intElms As IJElements, oIntrSecPoint As IJPoint
    Dim sX As Double, sY As Double, sZ As Double, eX As Double, eY As Double, eZ As Double
    Dim px As Double, py As Double, pz As Double, dist1 As Double, dist2 As Double
    Dim code As Geom3dIntersectConstants
    Dim bIntersects As Boolean, CutEnd As SPSMemberAxisPortIndex

    Set oSurf1 = oSuppingSurf
    
    '#79825 Fix Assert on trying to find intersection with surface. Whenever we try to find intersection with a surface from another component
    'we will use facelet's intersect method which overrides the intersect of IJSurface. That guy will assert
    'if a geometry implements IJLine does not implement IJGLine. SPSMemberPrismatic does not support IJGLine, but axis does
    Set oLine = oPart.Axis
    
    oLine.GetStartPoint sX, sY, sZ
    oLine.GetEndPoint eX, eY, eZ
    bIntersects = False
    
    oSurf1.Intersect oLine, intElms, code
    If Not intElms Is Nothing Then
        If intElms.count > 0 Then
            bIntersects = True
            Set oIntrSecPoint = intElms.Item(1)
            oIntrSecPoint.GetPoint px, py, pz
            dist1 = Sqr((px - sX) * (px - sX) + (py - sY) * (py - sY) + (pz - sZ) * (pz - sZ))
            dist2 = Sqr((px - eX) * (px - eX) + (py - eY) * (py - eY) + (pz - eZ) * (pz - eZ))
            If dist1 <= dist2 Then
                CutEnd = SPSMemberAxisStart
            Else
                CutEnd = SPSMemberAxisEnd
            End If
        End If
    End If
    Set oLine = Nothing
    Set oSurf1 = Nothing
    
    If bIntersects = False Then
        ' member line doesn't intersect the cutback surface ; figure out based on surface-surface int
        bIntersects = IsOperandOperatorIntersect(oPart, oSuppingSurf, 1, False, CutEnd) ' TR#75015
        If bIntersects = False Then ' TR#75015 ; now check with range box
            bIntersects = IsOperandOperatorIntersect(oPart, oSuppingSurf, 1, True, CutEnd)
        End If
    End If
    
    If bIntersects = True Then
        GetCutbackEnd = CutEnd
    Else
        GoTo ErrorHandler
    End If
    Exit Function

ErrorHandler:
    GetCutbackEnd = SPSMemberAxisStart
    HandleError MODULE, METHOD
End Function
Public Function IsSurfaceTrimCutbackNeeded(oSuppedPart As ISPSMemberPartPrismatic, oSuppingSurf As IJSurface) As Boolean
    Const METHOD = "IsSurfaceTrimCutbackNeeded"
    On Error GoTo ErrorHandler

    Dim oSurf1 As IJSurface, oLine As IJLine, intElms As IJElements
    Dim code As Geom3dIntersectConstants, bIsNeeded As Boolean, dummy As SPSMemberAxisPortIndex
    Dim myobj As Object, myplane As IJPlane, px As Double, py As Double, pz As Double, nx As Double, ny As Double, nz As Double
    Dim oGeomFact As New GeometryFactory
    Dim IsPlanar As Boolean
    Dim CosAngle As Double, plnNormal As IJDVector, memDir As IJDVector ' TR#75015
    
    Set myobj = oSuppingSurf
    If TypeOf myobj Is IJPlane Then
        ' now create an infinite plane and check for intersection
        Set myplane = oSuppingSurf
        myplane.GetRootPoint px, py, pz
        myplane.GetNormal nx, ny, nz

        Set plnNormal = New DVector
        Set memDir = New DVector
        Set oLine = oSuppedPart
        plnNormal.Set nx, ny, nz
        plnNormal.length = 1#
        oLine.GetDirection px, py, pz
        memDir.Set px, py, pz
        memDir.length = 1#
        Set oLine = Nothing
        
        CosAngle = Abs(plnNormal.x * memDir.x + plnNormal.y * memDir.y + plnNormal.z * memDir.z)
        If Abs(CosAngle) <= 0.0001 Then ' plane normal and member dir are perpendicular
            Set plnNormal = Nothing
            Set memDir = Nothing
            Set myplane = Nothing
            Set myobj = Nothing
            IsSurfaceTrimCutbackNeeded = False
            Exit Function
        End If
        Set plnNormal = Nothing
        Set memDir = Nothing
        Set myplane = Nothing
'        Set oSurf1 = oGeomFact.Planes3d.CreateByPointNormal(Nothing, px, py, pz, nX, nY, nZ)
        ' TR#75015 ; check with the original plane itself to avoid highlighting of invalid planes
        
        ' now it is ensured that the plane is not parallel to the member axis
        Set oSurf1 = oSuppingSurf
        IsPlanar = True
    Else
        Set oSurf1 = oSuppingSurf
        IsPlanar = False
    End If
    Set myobj = Nothing
    Set oGeomFact = Nothing
    '#79825 Fix Assert on trying to find intersection with surface.
    Set oLine = oSuppedPart.Axis
    bIsNeeded = False
    oSurf1.Intersect oLine, intElms, code
    If Not intElms Is Nothing Then
        If intElms.count > 0 Then
            bIsNeeded = DoesNotCutBothEnds(oSuppedPart, intElms) ' True TR#76302
        Else
            bIsNeeded = False
        End If
    Else
        bIsNeeded = False
    End If
    Set oLine = Nothing
    
    If bIsNeeded = False Then ' now check for surface-surface intersection ' TR#75015; removed "And IsPlanar = False"
        bIsNeeded = IsOperandOperatorIntersect(oSuppedPart, oSuppingSurf, 2, False, dummy) ' TR#75015
        If bIsNeeded = False Then ' TR#75015 ; now check with range box
            bIsNeeded = IsOperandOperatorIntersect(oSuppedPart, oSuppingSurf, 2, True, dummy)
        End If
    End If
    Set oSurf1 = Nothing
    IsSurfaceTrimCutbackNeeded = bIsNeeded
    Exit Function

ErrorHandler:
    IsSurfaceTrimCutbackNeeded = False
    HandleError MODULE, METHOD
End Function

Public Sub CreateSurfaceTrimCutback(oSuppedPart As ISPSMemberPartPrismatic, WhichEnd As SPSMemberAxisPortIndex, oSuppingSurf As IJSurface, oAttribsCAO As IJDAttributes, pResourceManager As IUnknown, ByRef oCutbackSurf As Object) 'IJSurface) ' TR#75102 ; added WhichEnd
    Const METHOD = "CreateSurfaceTrimCutback"
    On Error GoTo ErrorHandler

    Dim oLine As IJLine
    Dim sX As Double, sY As Double, sZ As Double, eX As Double, eY As Double, eZ As Double
    Dim px As Double, py As Double, pz As Double, CutbackClr As Double
    Dim toVec As IJDVector, awayVec As IJDVector, oPos1 As IJDPosition, oPos2 As IJDPosition
    Dim oPlaneRoot As IJDPosition, oPlaneNormal As IJDVector
    Dim oSurfRoot(1 To 25) As IJDPosition, oSurfNormal(1 To 25) As IJDVector, u As Double, v As Double, i As Integer
    Dim SurfPoints(0 To 74) As Double, uNumPoints(0 To 4) As Double, vNumPoints As Long
    Dim uForm As Geom3dCurveFormConstants, vForm As Geom3dCurveFormConstants
    Dim pGeometryFactory As New GeometryFactory
    Dim InfLine(0 To 1) As IJDPosition
    Dim SquareEnd As Boolean, IsPlanar As Boolean, myobj As Object, oSuppingPlane As IJPlane
    Dim FlangeAngle As Double, WebAngle As Double
    
    Set toVec = New DVector
    Set awayVec = New DVector
    Set oPos1 = New DPosition
    Set oPos2 = New DPosition
    Set oPlaneRoot = New DPosition
    Set oPlaneNormal = New DVector

    Set InfLine(0) = New DPosition
    Set InfLine(1) = New DPosition
    CutbackClr = oAttribsCAO.CollectionOfAttributes("IJUASPSSurfaceTrim").Item("Offset").Value
    SquareEnd = oAttribsCAO.CollectionOfAttributes("IJUASPSSurfaceTrim").Item("SquaredEnd").Value
    Set myobj = oSuppingSurf
    If TypeOf myobj Is IJPlane Then
        IsPlanar = True
        Set oSuppingPlane = oSuppingSurf
    Else
        IsPlanar = False
    End If
    Set myobj = Nothing
    
    If IsPlanar Then
        oSuppingPlane.GetRootPoint px, py, pz
        oPlaneRoot.Set px, py, pz
        oSuppingPlane.GetNormal px, py, pz
        oPlaneNormal.Set px, py, pz
        oPlaneNormal.length = 1#
    Else ' surface code
        ' do nothing; old code removed
    End If
    
    Set oLine = oSuppedPart
    oLine.GetStartPoint sX, sY, sZ
    oLine.GetEndPoint eX, eY, eZ
    
'    WhichEnd = GetCutbackEnd(oSuppedPart, oSuppingSurf) TR#75102
    If WhichEnd = SPSMemberAxisStart Then
        ' member start is being cutback
        oPos1.Set eX, eY, eZ
    Else
        ' member end is being cutback
        oPos1.Set sX, sY, sZ
    End If
    
    If IsPlanar Then
        InfLine(0).Set oPlaneRoot.x, oPlaneRoot.y, oPlaneRoot.z
        InfLine(1).Set oPlaneRoot.x + LARGE_EDGE * oPlaneNormal.x, oPlaneRoot.y + LARGE_EDGE * oPlaneNormal.y, oPlaneRoot.z + LARGE_EDGE * oPlaneNormal.z
        ProjectPtOnLine oPos1, InfLine, oPos2
        toVec.Set oPos2.x - oPlaneRoot.x, oPos2.y - oPlaneRoot.y, oPos2.z - oPlaneRoot.z
        toVec.length = 1#
        awayVec.Set -1 * toVec.x, -1 * toVec.y, -1 * toVec.z
    
        ' now move the cutting plane such that the clearance is satisfied
        px = oPlaneRoot.x + CutbackClr * toVec.x
        py = oPlaneRoot.y + CutbackClr * toVec.y
        pz = oPlaneRoot.z + CutbackClr * toVec.z
        oPlaneRoot.Set px, py, pz
    Else ' surface code
        ' do nothing; old code removed
    End If
    
    ' TR#77632 ; moved following lines from inside 'if' block
    Dim oIJGeometryMisc As IJGeometryMisc
    Dim myplane As IJPlane
    Dim tmpSheet As Object
    Dim oIJDSurfacedefintion As IJDSurfaceDefinition
    Set oIJDSurfacedefintion = New DGeomOpsSurfDef
    Set oIJGeometryMisc = New DGeomOpsMisc
    If IsPlanar Then
        If oCutbackSurf Is Nothing Then
            Set myplane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, oPlaneRoot.x, oPlaneRoot.y, oPlaneRoot.z, oPlaneNormal.x, oPlaneNormal.y, oPlaneNormal.z) ' TR#77632 ; create a transient plane
            oIJGeometryMisc.CreateModelGeometryFromGType pResourceManager, myplane, Nothing, oCutbackSurf ' TR#77632
            Set myplane = Nothing
        Else
            Set myplane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, oPlaneRoot.x, oPlaneRoot.y, oPlaneRoot.z, oPlaneNormal.x, oPlaneNormal.y, oPlaneNormal.z) ' TR#77632 ; create a transient plane' oCutbackSurf
            oIJGeometryMisc.ModifyModelGeometryFromGType myplane, oCutbackSurf ' TR#77632
            Set myplane = Nothing
        End If
    Else
'        If oCutbackSurf Is Nothing Then
'            Set oCutbackSurf = oSuppingSurf ' pGeometryFactory.BSplineSurfaces3d.CreateByFitSurface(pResourceManager, vNumPoints, uNumPoints, SurfPoints, 4, 4, uForm, vForm)

        ' convert to sheet
        oIJGeometryMisc.CreateModelGeometryFromGType Nothing, oSuppingSurf, Nothing, tmpSheet
        If oCutbackSurf Is Nothing Then
            oIJDSurfacedefintion.PlaceSurfaceByOffset pResourceManager, tmpSheet, CutbackClr, Nothing, oCutbackSurf, Nothing
        Else
            oIJDSurfacedefintion.ModifySurfaceByOffset pResourceManager, tmpSheet, CutbackClr, oCutbackSurf, Nothing
        End If
        
'        End If
    End If
    Set oIJGeometryMisc = Nothing ' TR#77632 moved from inside 'else' block
    Set oIJDSurfacedefintion = Nothing

    ' now set all the outputs for the surface cutback
    ' call the new function from feature services for planar trim TR 80543
    If SquareEnd Then
        Dim oFeatureServices As ISPSMemberFeatureServices
        Dim oMemberFactory As SPSMemberFactory
        Dim outPosX As Double
        Dim outPosY As Double
        Dim outPosZ As Double
        Dim outNorX As Double
        Dim outNorY As Double
        Dim outNorZ As Double
        Dim allOK As Boolean
        allOK = True
        
        Set oMemberFactory = New SPSMemberFactory
        Set oFeatureServices = oMemberFactory.CreateMemberFeatureServices
        oFeatureServices.ComputePlanarTrim oSuppedPart, WhichEnd, oSuppingSurf, CutbackClr, outPosX, outPosY, outPosZ, outNorX, outNorY, outNorZ, allOK
        
        If allOK Then
            ' Now build the surface for cutback
            'Dim oIJGeometryMisc As IJGeometryMisc
            Set oIJGeometryMisc = New DGeomOpsMisc
            
            If oCutbackSurf Is Nothing Then
                Set myplane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, outPosX, outPosY, outPosZ, outNorX, outNorY, outNorZ)
                oIJGeometryMisc.CreateModelGeometryFromGType pResourceManager, myplane, Nothing, oCutbackSurf
                Set myplane = Nothing
            ElseIf IsPlanar = True Then
                Set myplane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, outPosX, outPosY, outPosZ, outNorX, outNorY, outNorZ)
                oIJGeometryMisc.ModifyModelGeometryFromGType myplane, oCutbackSurf
                Set myplane = Nothing
            Else ' non-planar surface
                Dim tmpPlane As IJPlane
    
                Set tmpPlane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, outPosX, outPosY, outPosZ, outNorX, outNorY, outNorZ)
                oIJGeometryMisc.ModifyModelGeometryFromGType tmpPlane, oCutbackSurf
                Set tmpPlane = Nothing
            End If
            Set oIJGeometryMisc = Nothing
        End If
        Set oFeatureServices = Nothing
        Set oMemberFactory = Nothing
        
    Else
        ComputeSurfaceTrimShapeAndOutputs oSuppedPart, SquareEnd, WhichEnd, pResourceManager, oCutbackSurf, FlangeAngle, WebAngle  ' TR#75014 ; passing WhichEnd argument
    End If

    oAttribsCAO.CollectionOfAttributes("IJUASPSSurfaceTrim").Item("FlangeAngle").Value = FlangeAngle
    oAttribsCAO.CollectionOfAttributes("IJUASPSSurfaceTrim").Item("WebAngle").Value = WebAngle
    oAttribsCAO.CollectionOfAttributes("IJUASPSSurfaceTrim").Item("TrimEnd").Value = WhichEnd
    
    Set oSuppingPlane = Nothing
    Set pGeometryFactory = Nothing
    Set oPlaneRoot = Nothing
    Set oPlaneNormal = Nothing
    Set InfLine(0) = Nothing
    Set InfLine(1) = Nothing
    Set toVec = Nothing
    Set awayVec = Nothing
    Set oPos1 = Nothing
    Set oPos2 = Nothing
    Set oLine = Nothing
    For i = 1 To 25
        Set oSurfRoot(i) = Nothing
        Set oSurfNormal(i) = Nothing
    Next i
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Public Function IsOperandOperatorIntersect(oSuppedPart As ISPSMemberPartPrismatic, oSuppingSurf As IJSurface, ByVal Mode As Integer, ByVal Bounding As Boolean, ByRef CutEnd As SPSMemberAxisPortIndex) As Boolean ' TR#75015 ; added Bounding flag to validate against range box
    Const METHOD = "IsOperandOperatorIntersect"
    On Error GoTo ErrorHandler

    ' checks for plane to plane intersection instead of member axis and supporting plane intersection
    Dim bIsNeeded As Boolean
    Dim oSurf1 As IJSurface, oSurf2 As IJSurface, intElms As IJElements
    Dim code As Geom3dIntersectConstants, IntLine As Line3d, i As Integer, nextIdx As Integer
    Dim oStartPos() As IJDPosition, oEndPos() As IJDPosition
    Dim pGeometryFactory As New GeometryFactory
    Dim PlanePoints() As Double
    Dim oLine As IJLine, oCurve As IJCurve, dist1 As Double, dist2 As Double
    Dim par As Double, px As Double, py As Double, pz As Double, sX As Double, sY As Double, sZ As Double, eX As Double, eY As Double, eZ As Double
    Dim PtCount As Integer
    
    ' Mode = 1 ; called from corner-cope ; check for end faces also
    ' Mode = 2 ; called from surface-trim; check for only side faces
    bIsNeeded = False
    CutEnd = SPSMemberAxisStart
    Set oSurf1 = oSuppingSurf

    GetCrossSectionPoints oSuppedPart, SPSMemberAxisStart, Bounding, PtCount, oStartPos ' TR#75015
    GetCrossSectionPoints oSuppedPart, SPSMemberAxisEnd, Bounding, PtCount, oEndPos ' TR#75015
    ReDim PlanePoints(0 To PtCount * 3) As Double
    Set oLine = oSuppedPart
    oLine.GetStartPoint sX, sY, sZ
    oLine.GetEndPoint eX, eY, eZ
    Set oLine = Nothing
    If bIsNeeded = False Then
        ' now check with side faces
        For i = 0 To PtCount - 1
            nextIdx = (i + 1) Mod 4
            PlanePoints(0) = oStartPos(i).x
            PlanePoints(1) = oStartPos(i).y
            PlanePoints(2) = oStartPos(i).z
            PlanePoints(3) = oStartPos(nextIdx).x
            PlanePoints(4) = oStartPos(nextIdx).y
            PlanePoints(5) = oStartPos(nextIdx).z
            PlanePoints(6) = oEndPos(nextIdx).x
            PlanePoints(7) = oEndPos(nextIdx).y
            PlanePoints(8) = oEndPos(nextIdx).z
            PlanePoints(9) = oEndPos(i).x
            PlanePoints(10) = oEndPos(i).y
            PlanePoints(11) = oEndPos(i).z
            Set oSurf2 = pGeometryFactory.Planes3d.CreateByPoints(Nothing, 4, PlanePoints)
            oSurf1.Intersect oSurf2, intElms, code
            Set oSurf2 = Nothing
            If Not intElms Is Nothing Then
                If intElms.count > 0 Then
                    bIsNeeded = DoesNotCutBothEnds(oSuppedPart, intElms) ' True  TR#76302
                    If bIsNeeded = True Then
                        Set oCurve = intElms.Item(1)
                        oCurve.ParameterFRatio 0.5, par
                        oCurve.Position par, px, py, pz
                        dist1 = Sqr((px - sX) * (px - sX) + (py - sY) * (py - sY) + (pz - sZ) * (pz - sZ))
                        dist2 = Sqr((px - eX) * (px - eX) + (py - eY) * (py - eY) + (pz - eZ) * (pz - eZ))
                        If dist1 <= dist2 Then
                            CutEnd = SPSMemberAxisStart
                        Else
                            CutEnd = SPSMemberAxisEnd
                        End If
                        Set oCurve = Nothing
                        Exit For
                    End If
                End If
            End If
        Next i
    End If
    
    If bIsNeeded = False And Mode = 1 Then
        ' check for intersection with start endshape
        For i = 0 To PtCount - 1
            PlanePoints(i * 3) = oStartPos(i).x
            PlanePoints(i * 3 + 1) = oStartPos(i).y
            PlanePoints(i * 3 + 2) = oStartPos(i).z
        Next i
        Set oSurf2 = pGeometryFactory.Planes3d.CreateByPoints(Nothing, PtCount, PlanePoints)
        oSurf1.Intersect oSurf2, intElms, code
        Set oSurf2 = Nothing
        If Not intElms Is Nothing Then
            If intElms.count > 0 Then
                bIsNeeded = DoesNotCutBothEnds(oSuppedPart, intElms) ' True TR#76302
                If bIsNeeded = True Then
                    Set oCurve = intElms.Item(1)
                    oCurve.ParameterFRatio 0.5, par
                    oCurve.Position par, px, py, pz
                    dist1 = Sqr((px - sX) * (px - sX) + (py - sY) * (py - sY) + (pz - sZ) * (pz - sZ))
                    dist2 = Sqr((px - eX) * (px - eX) + (py - eY) * (py - eY) + (pz - eZ) * (pz - eZ))
                    If dist1 <= dist2 Then
                        CutEnd = SPSMemberAxisStart
                    Else
                        CutEnd = SPSMemberAxisEnd
                    End If
                    Set oCurve = Nothing
                End If
            End If
        End If
    End If
    
    If bIsNeeded = False And Mode = 1 Then
        ' now check with end endshape
        For i = 0 To PtCount - 1
            PlanePoints(i * 3) = oEndPos(i).x
            PlanePoints(i * 3 + 1) = oEndPos(i).y
            PlanePoints(i * 3 + 2) = oEndPos(i).z
        Next i
        Set oSurf2 = pGeometryFactory.Planes3d.CreateByPoints(Nothing, PtCount, PlanePoints)
        oSurf1.Intersect oSurf2, intElms, code
        Set oSurf2 = Nothing
        If Not intElms Is Nothing Then
            If intElms.count > 0 Then
                bIsNeeded = DoesNotCutBothEnds(oSuppedPart, intElms) ' True TR#76302
                If bIsNeeded = True Then
                    Set oCurve = intElms.Item(1)
                    oCurve.ParameterFRatio 0.5, par
                    oCurve.Position par, px, py, pz
                    dist1 = Sqr((px - sX) * (px - sX) + (py - sY) * (py - sY) + (pz - sZ) * (pz - sZ))
                    dist2 = Sqr((px - eX) * (px - eX) + (py - eY) * (py - eY) + (pz - eZ) * (pz - eZ))
                    If dist1 <= dist2 Then
                        CutEnd = SPSMemberAxisStart
                    Else
                        CutEnd = SPSMemberAxisEnd
                    End If
                    Set oCurve = Nothing
                End If
            End If
        End If
    End If

    Set oSurf1 = Nothing
    IsOperandOperatorIntersect = bIsNeeded
    Exit Function

ErrorHandler:
    IsOperandOperatorIntersect = False
    CutEnd = SPSMemberAxisStart
    HandleError MODULE, METHOD
End Function

'Private Sub CreateCopeCuttingShapeAndComputeOutputs(oSuppedPart As ISPSMemberPartPrismatic, ByVal PlaneIntPt As IJDPosition, ByVal Shp1Vec As IJDVector, ByVal Shp2Vec As IJDVector, ByRef SideClr As Double, ByVal Clr1Vec As IJDVector, ByRef InsideClr As Double, ByVal Clr2Vec As IJDVector, ByVal RoundIncrement As Double, ByVal SquaredEnd As Boolean, ByVal PlanesConcave As Boolean, ByVal Radius As Double, ByVal RadiusType As Integer, ByRef curveElms As IJElements, ByRef CopeLength As Double, ByRef CopeDepth As Double, ByVal CopeType As Integer, ByRef CopeSide As Integer, ByRef CopeEdge As Integer, ByRef WhichEnd As SPSMemberAxisPortIndex)
'    Const METHOD = "CreateCopeCuttingShapeAndComputeOutputs"
'    On Error GoTo ErrorHandler
'    Dim px As Double, py As Double, pz As Double
'    Dim Circular As Boolean, i As Integer
'    Dim oCurve1 As IJCurve, oCurve2 As IJCurve
'    Dim pGeometryFactory As New GeometryFactory
'    Dim idx As Integer
'    Dim TmpVec1 As IJDVector, TmpVec2 As IJDVector
'    Dim NumInts As Long, IntPoints() As Double, overlap As Long
'    Dim ProjStartPos As IJDPosition, ProjEndPos As IJDPosition, oStartPos(0 To 3) As IJDPosition, oEndPos(0 To 3) As IJDPosition
'    Dim minx As Double, miny As Double, minz As Double, maxx As Double, maxy As Double, maxz As Double
'    Dim oCmplx1 As ComplexString3d, oCmplx2 As ComplexString3d
'    Dim MemStartPos As IJDPosition, MemEndPos As IJDPosition, oLine As IJLine, TmpPos As IJDPosition, TmpPos2 As IJDPosition
'    Dim ProjLength As Double, ProjWidth As Double, TmpVec3 As IJDVector, TmpVec4 As IJDVector
'    Dim memElms As IJElements, LinePos(0 To 1) As IJDPosition
'    Dim MinLenDist As Double, MaxLenDist As Double, MinWidDist As Double, MaxWidDist As Double, TmpDist As Double, TmpDist2 As Double
'    Dim TmpVec5 As IJDVector, TmpVec6 As IJDVector, FirstIter As Boolean
'    Dim ProjIntSet(0 To 1) As Boolean, MinIntPt As IJDPosition, MaxIntPt As IJDPosition
'    Dim NewCopeLength As Double, NewCopeDepth As Double, dblResult As Double, intResult As Integer
'    Dim DeltaCL As Double, DeltaCD As Double, DeltaSC As Double, DeltaIC As Double, ICComp As Double, IsAtBottom As Boolean, IsAtLeft As Boolean
'    Dim IsAtStart As Boolean, MinWidOther As Double, MaxWidOther As Double, RefPt As IJDPosition, TmpPos3 As IJDPosition
'    Dim DrawMode As Integer, DepthCut As Boolean
'
'    For i = 0 To 3
'        Set oStartPos(i) = New DPosition
'        Set oEndPos(i) = New DPosition
'    Next i
'    DepthCut = False
'    Set RefPt = New DPosition
'    Set TmpPos3 = New DPosition
'    Set MinIntPt = New DPosition
'    Set MaxIntPt = New DPosition
'    Set TmpVec1 = New DVector
'    Set TmpVec2 = New DVector
'    Set TmpVec3 = New DVector
'    Set TmpVec4 = New DVector
'    Set TmpVec5 = New DVector
'    Set TmpVec6 = New DVector
'    Set MemStartPos = New DPosition
'    Set MemEndPos = New DPosition
'    Set ProjStartPos = New DPosition
'    Set ProjEndPos = New DPosition
'    Set TmpPos = New DPosition
'    Set TmpPos2 = New DPosition
'    Set LinePos(0) = New DPosition
'    Set LinePos(1) = New DPosition
'
'    Set curveElms = New JObjectCollection
'
'    ' now get the rectangular outline of the member part and then get the intersection points
'    Dim oProfile As IJCrossSection
'    Set oProfile = oSuppedPart.CrossSection.Definition
'    Select Case oProfile.Type
'        Case "HSSC", "CS", "PIPE"
'            Circular = True
'        Case Else
'            Circular = False
'    End Select
'    Set oProfile = Nothing
'
'    Set oLine = oSuppedPart
'    oLine.GetStartPoint px, py, pz
'    MemStartPos.Set px, py, pz
'    oLine.GetEndPoint px, py, pz
'    MemEndPos.Set px, py, pz
'    Set oLine = Nothing
'    Set TmpVec1 = Shp1Vec.Cross(Shp2Vec)
'    TmpVec1.length = 1#
'    Set ProjStartPos = ProjectPosToPlane(MemStartPos, PlaneIntPt, TmpVec1)
'    Set ProjEndPos = ProjectPosToPlane(MemEndPos, PlaneIntPt, TmpVec1)
'    TmpVec2.Set ProjEndPos.x - ProjStartPos.x, ProjEndPos.y - ProjStartPos.y, ProjEndPos.z - ProjStartPos.z
'    TmpVec2.length = 1#
'    Set TmpVec3 = TmpVec1.Cross(TmpVec2)
'    TmpVec3.length = 1#
'    TmpDist = Sqr((PlaneIntPt.x - ProjStartPos.x) * (PlaneIntPt.x - ProjStartPos.x) + (PlaneIntPt.y - ProjStartPos.y) * (PlaneIntPt.y - ProjStartPos.y) + (PlaneIntPt.z - ProjStartPos.z) * (PlaneIntPt.z - ProjStartPos.z))
'    TmpDist2 = Sqr((PlaneIntPt.x - ProjEndPos.x) * (PlaneIntPt.x - ProjEndPos.x) + (PlaneIntPt.y - ProjEndPos.y) * (PlaneIntPt.y - ProjEndPos.y) + (PlaneIntPt.z - ProjEndPos.z) * (PlaneIntPt.z - ProjEndPos.z))
'    If TmpDist <= TmpDist2 Then
'        WhichEnd = SPSMemberAxisStart
'    Else
'        WhichEnd = SPSMemberAxisEnd
'    End If
'
'    If Circular = False Then
'        GetBoundingRectangleForMemberEnd oSuppedPart, SPSMemberAxisStart, oStartPos
'        GetBoundingRectangleForMemberEnd oSuppedPart, SPSMemberAxisEnd, oEndPos
'        For i = 0 To 3
'            Set TmpPos = ProjectPosToPlane(oStartPos(i), PlaneIntPt, TmpVec1)
'            TmpVec4.Set TmpPos.x - ProjStartPos.x, TmpPos.y - ProjStartPos.y, TmpPos.z - ProjStartPos.z
'            If TmpVec4.length > 0.0001 Then
'                TmpDist = TmpVec4.Dot(TmpVec2)
'            Else
'                TmpDist = 0#
'            End If
'            If MinLenDist > TmpDist Or i = 0 Then MinLenDist = TmpDist
'            If MaxLenDist < TmpDist Or i = 0 Then MaxLenDist = TmpDist
'            If TmpVec4.length > 0.0001 Then
'                TmpDist = TmpVec4.Dot(TmpVec3)
'            Else
'                TmpDist = 0#
'            End If
'            If MinWidDist > TmpDist Or i = 0 Then MinWidDist = TmpDist
'            If MaxWidDist < TmpDist Or i = 0 Then MaxWidDist = TmpDist
'
'            Set TmpPos = ProjectPosToPlane(oEndPos(i), PlaneIntPt, TmpVec1)
'            TmpVec4.Set TmpPos.x - ProjStartPos.x, TmpPos.y - ProjStartPos.y, TmpPos.z - ProjStartPos.z
'            If TmpVec4.length > 0.0001 Then
'                TmpDist = TmpVec4.Dot(TmpVec2)
'            Else
'                TmpDist = 0#
'            End If
'            If MinLenDist > TmpDist Then MinLenDist = TmpDist
'            If MaxLenDist < TmpDist Then MaxLenDist = TmpDist
'            If TmpVec4.length > 0.0001 Then
'                TmpDist = TmpVec4.Dot(TmpVec3)
'            Else
'                TmpDist = 0#
'            End If
'            If MinWidDist > TmpDist Then MinWidDist = TmpDist
'            If MaxWidDist < TmpDist Then MaxWidDist = TmpDist
'        Next i
'    Else ' circular section ; can get the projected range box directly
'        Dim oSectionAttrbs As IJDAttributes, diameter As Double
'        Set oSectionAttrbs = oSuppedPart.CrossSection.Definition
'        diameter = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
'
'        MinLenDist = 0
'        MaxLenDist = ProjEndPos.DistPt(ProjStartPos)
'        MinWidDist = -0.5 * diameter
'        MaxWidDist = 0.5 * diameter
'    End If
'
'    TmpPos.Set ProjStartPos.x + MinWidDist * TmpVec3.x, ProjStartPos.y + MinWidDist * TmpVec3.y, ProjStartPos.z + MinWidDist * TmpVec3.z
'    minx = TmpPos.x + MinLenDist * TmpVec2.x
'    miny = TmpPos.y + MinLenDist * TmpVec2.y
'    minz = TmpPos.z + MinLenDist * TmpVec2.z
'    TmpPos.Set ProjStartPos.x + MaxWidDist * TmpVec3.x, ProjStartPos.y + MaxWidDist * TmpVec3.y, ProjStartPos.z + MaxWidDist * TmpVec3.z
'    maxx = TmpPos.x + MaxLenDist * TmpVec2.x
'    maxy = TmpPos.y + MaxLenDist * TmpVec2.y
'    maxz = TmpPos.z + MaxLenDist * TmpVec2.z
'
'    TmpVec4.Set maxx - minx, maxy - miny, maxz - minz
'    ProjLength = Abs(TmpVec4.Dot(TmpVec2))
'    ProjWidth = Sqr(TmpVec4.length * TmpVec4.length - ProjLength * ProjLength)
'    TmpPos.Set maxx, maxy, maxz
'    LinePos(0).Set minx, miny, minz
'    LinePos(1).Set minx + ProjWidth * TmpVec3.x, miny + ProjWidth * TmpVec3.y, minz + ProjWidth * TmpVec3.z
'    ProjectPtOnLine TmpPos, LinePos, TmpPos2
'    TmpVec4.Set TmpPos2.x - minx, TmpPos2.y - miny, TmpPos2.z - minz
'    TmpVec4.length = 1#
'    LinePos(1).Set minx + ProjLength * TmpVec2.x, miny + ProjLength * TmpVec2.y, minz + ProjLength * TmpVec2.z
'    ProjectPtOnLine TmpPos, LinePos, TmpPos2
'    TmpVec3.Set TmpPos2.x - minx, TmpPos2.y - miny, TmpPos2.z - minz
'    TmpVec3.length = 1#
'
'    LinePos(0).Set oStartPos(0).x, oStartPos(0).y, oStartPos(0).z
'    LinePos(1).Set oEndPos(0).x, oEndPos(0).y, oEndPos(0).z
'    TmpPos.Set minx, miny, minz
'    ProjectPtOnLine TmpPos, LinePos, TmpPos2
'    TmpDist = TmpPos2.DistPt(oStartPos(0))
'    TmpDist2 = TmpPos2.DistPt(oEndPos(0))
'    If TmpDist <= TmpDist2 Then
'        ' pt minx, miny, minz is at start
'        IsAtStart = True
'    Else
'        ' pt minx, miny, minz is at end
'        IsAtStart = False
'    End If
'    If CopeType = 1 Then ' webcope
'        LinePos(0).Set oStartPos(0).x, oStartPos(0).y, oStartPos(0).z
'        LinePos(1).Set oStartPos(3).x, oStartPos(3).y, oStartPos(3).z
'        TmpPos.Set minx, miny, minz
'        ProjectPtOnLine TmpPos, LinePos, TmpPos2
'        TmpDist = TmpPos2.DistPt(oStartPos(0))
'        TmpDist2 = TmpPos2.DistPt(oStartPos(3))
'        If TmpDist <= TmpDist2 Then
'            ' pt minx, miny, minz is at bottom
'            IsAtBottom = True
'        Else
'            ' pt minx, miny, minz is at bottom
'            IsAtBottom = False
'        End If
'    ElseIf CopeType = 2 And Circular = True Then ' flangecope and circular cross section
'        LinePos(0).Set oStartPos(0).x, oStartPos(0).y, oStartPos(0).z
'        LinePos(1).Set oStartPos(1).x, oStartPos(1).y, oStartPos(1).z
'        TmpPos.Set minx, miny, minz
'        ProjectPtOnLine TmpPos, LinePos, TmpPos2
'        TmpDist = TmpPos2.DistPt(oStartPos(0))
'        TmpDist2 = TmpPos2.DistPt(oStartPos(1))
'        If TmpDist <= TmpDist2 Then
'            ' pt minx, miny, minz is at left
'            IsAtLeft = True
'        Else
'            ' pt minx, miny, minz is at right
'            IsAtLeft = False
'        End If
'    ElseIf CopeType = 0 Then ' generic cope ; no need to populate copeside
'        CopeSide = 0
'        CopeEdge = 0
'    End If
'
'    ' now construct a member shape using min and max
'    ' tmpvec3 is along projected length, tmpvec4 is along projected width
'    Set memElms = New JObjectCollection
'    Set oLine = Nothing
'    Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, minx, miny, minz, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'    memElms.Add oLine
'    Set oLine = Nothing
'    Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, minx + ProjLength * TmpVec3.x, miny + ProjLength * TmpVec3.y, minz + ProjLength * TmpVec3.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'    memElms.Add oLine
'    Set oLine = Nothing
'    Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, maxx, maxy, maxz, -1# * TmpVec3.x, -1# * TmpVec3.y, -1# * TmpVec3.z, ProjLength)
'    memElms.Add oLine
'    Set oLine = Nothing
'    Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, minx + ProjWidth * TmpVec4.x, miny + ProjWidth * TmpVec4.y, minz + ProjWidth * TmpVec4.z, -1# * TmpVec4.x, -1# * TmpVec4.y, -1# * TmpVec4.z, ProjWidth)
'    memElms.Add oLine
'    Set oLine = Nothing
'    Set oCmplx2 = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, memElms)
'
'    ' now draw the cope cutting shape taking into consideration clearances etc
'    FirstIter = True
'    DrawMode = 1
'DRAWSHAPE:
'    DrawCornerCopeShape DrawMode, PlanesConcave, Radius, RadiusType, PlaneIntPt, Nothing, Shp1Vec, Shp2Vec, SideClr, Clr1Vec, InsideClr, Clr2Vec, curveElms
'
'    Set oCmplx1 = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
'    GetIntersectionPointsOfCurves oCmplx1, oCmplx2, NumInts, IntPoints
'    Set oCmplx1 = Nothing
'
'    MinLenDist = ProjLength
'    MaxLenDist = 0
'    MinWidDist = ProjWidth
'    MaxWidDist = 0
'    ProjIntSet(0) = False
'    ProjIntSet(1) = False
'    For i = 1 To NumInts
'        TmpPos.Set IntPoints((i - 1) * 3), IntPoints((i - 1) * 3 + 1), IntPoints((i - 1) * 3 + 2)
'
'        ' first in projlength direction
'        LinePos(0).Set minx, miny, minz
'        LinePos(1).Set minx + ProjLength * TmpVec3.x, miny + ProjLength * TmpVec3.y, minz + ProjLength * TmpVec3.z
'        ProjectPtOnLine TmpPos, LinePos, TmpPos2
'        TmpDist = Sqr((TmpPos2.x - minx) * (TmpPos2.x - minx) + (TmpPos2.y - miny) * (TmpPos2.y - miny) + (TmpPos2.z - minz) * (TmpPos2.z - minz))
'        If TmpDist <= 0.0001 Then
'            ProjIntSet(0) = True
'            MinIntPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'        ElseIf Abs(TmpDist - ProjLength) <= 0.0001 Then
'            ProjIntSet(1) = True
'            MaxIntPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'        End If
'        If MinLenDist > TmpDist Then MinLenDist = TmpDist
'        If MaxLenDist < TmpDist Then MaxLenDist = TmpDist
'
'        ' now in projwidth direction
'        LinePos(1).Set minx + ProjWidth * TmpVec4.x, miny + ProjWidth * TmpVec4.y, minz + ProjWidth * TmpVec4.z
'        ProjectPtOnLine TmpPos, LinePos, TmpPos2
'        TmpDist = Sqr((TmpPos2.x - minx) * (TmpPos2.x - minx) + (TmpPos2.y - miny) * (TmpPos2.y - miny) + (TmpPos2.z - minz) * (TmpPos2.z - minz))
'        If MinWidDist > TmpDist Then MinWidDist = TmpDist
'        If MaxWidDist < TmpDist Then MaxWidDist = TmpDist
'    Next i
'    ' now do the same for PlaneIntPoint in ProjLength dir
'    LinePos(0).Set minx, miny, minz
'    LinePos(1).Set minx + ProjLength * TmpVec3.x, miny + ProjLength * TmpVec3.y, minz + ProjLength * TmpVec3.z
'    ProjectPtOnLine PlaneIntPt, LinePos, TmpPos2
'    TmpDist = Sqr((TmpPos2.x - minx) * (TmpPos2.x - minx) + (TmpPos2.y - miny) * (TmpPos2.y - miny) + (TmpPos2.z - minz) * (TmpPos2.z - minz))
'    If MinLenDist > TmpDist Then MinLenDist = TmpDist
'    If MaxLenDist < TmpDist Then MaxLenDist = TmpDist
'    LinePos(1).Set minx + ProjWidth * TmpVec4.x, miny + ProjWidth * TmpVec4.y, minz + ProjWidth * TmpVec4.z
'    ProjectPtOnLine PlaneIntPt, LinePos, TmpPos2
'    TmpDist = Sqr((TmpPos2.x - minx) * (TmpPos2.x - minx) + (TmpPos2.y - miny) * (TmpPos2.y - miny) + (TmpPos2.z - minz) * (TmpPos2.z - minz))
'    If MinWidDist > TmpDist Then MinWidDist = TmpDist
'    If MaxWidDist < TmpDist Then MaxWidDist = TmpDist
'
'    If Abs(MaxLenDist - MinLenDist) <= 0.0001 Then
'        CopeLength = MaxLenDist
'    Else
'        CopeLength = MaxLenDist - MinLenDist
'    End If
'    If Abs(MaxWidDist - MinWidDist) <= 0.0001 Then
'        CopeDepth = MaxWidDist
'    Else
'        CopeDepth = MaxWidDist - MinWidDist
'    End If
'
'    If CopeLength < ProjLength And CopeDepth < ProjWidth Then
'        If Abs(MinWidDist) <= 0.0001 Then
'            If CopeType = 1 Then
'                If IsAtBottom = True Then
'                    CopeEdge = 513 ' bottom
'                Else
'                    CopeEdge = 514 ' top
'                End If
'            ElseIf CopeType = 2 And Circular = True Then
'                    If IsAtLeft = True Then
'                        CopeSide = 1 ' left
'                        CopeEdge = 514 ' top
'                    Else
'                        CopeSide = 2 ' right
'                        CopeEdge = 514 ' top
'                    End If
'            End If
'            If Abs(MinLenDist) <= 0.0001 Then
'                ' case 1 in diagram
'                If IsAtStart = True Then
'                    WhichEnd = SPSMemberAxisStart
'                Else
'                    WhichEnd = SPSMemberAxisEnd
'                End If
'                Set oCurve1 = Nothing
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MaxLenDist * TmpVec3.x, miny + MaxLenDist * TmpVec3.y, minz + MaxLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve1 = oLine
'                Set oLine = Nothing
'                TmpPos2.Set minx + MaxWidDist * TmpVec4.x, miny + MaxWidDist * TmpVec4.y, minz + MaxWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos2.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                    TmpVec5.Set -TmpVec3.x, -TmpVec3.y, -TmpVec3.z
'                    TmpVec6.Set -TmpVec4.x, -TmpVec4.y, -TmpVec4.z
'                    If SquaredEnd = True Then
'                        DrawCornerCopeShape 1, False, Radius, RadiusType, TmpPos2, Nothing, TmpVec5, TmpVec6, 0, Nothing, 0, Nothing, curveElms
'                        DrawMode = 1
'                        PlanesConcave = False
'                        PlaneIntPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                        Shp1Vec.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                        Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                        Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                        Clr2Vec.Set -Shp1Vec.x, -Shp1Vec.y, -Shp1Vec.z
'                        SideClr = 0#
'                        InsideClr = 0#
'                    End If
'                End If
'            ElseIf Abs(MaxLenDist - ProjLength) <= 0.0001 Then
'                ' case 2 of diagram
'                If IsAtStart = True Then
'                    WhichEnd = SPSMemberAxisEnd
'                Else
'                    WhichEnd = SPSMemberAxisStart
'                End If
'                Set oCurve1 = Nothing
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MinLenDist * TmpVec3.x, miny + MinLenDist * TmpVec3.y, minz + MinLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve1 = oLine
'                Set oLine = Nothing
'                TmpPos2.Set minx + MaxWidDist * TmpVec4.x, miny + MaxWidDist * TmpVec4.y, minz + MaxWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos2.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                    TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                    TmpVec6.Set -TmpVec4.x, -TmpVec4.y, -TmpVec4.z
'                    If SquaredEnd = True Then
'                        DrawCornerCopeShape 1, False, Radius, RadiusType, TmpPos2, Nothing, TmpVec5, TmpVec6, 0, Nothing, 0, Nothing, curveElms
'                        DrawMode = 1
'                        PlanesConcave = False
'                        PlaneIntPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                        Shp1Vec.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                        Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                        Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                        Clr2Vec.Set -Shp1Vec.x, -Shp1Vec.y, -Shp1Vec.z
'                        SideClr = 0#
'                        InsideClr = 0#
'                    End If
'                End If
'            Else
'                ' case 3 of diagram ; minlendist and maxlendist are between the extremes
'                Set oCurve1 = Nothing
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MaxWidDist * TmpVec4.x, miny + MaxWidDist * TmpVec4.y, minz + MaxWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve1 = oLine
'                Set oLine = Nothing
'                TmpPos2.Set minx + MinLenDist * TmpVec3.x, miny + MinLenDist * TmpVec3.y, minz + MinLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MaxLenDist * TmpVec3.x, miny + MaxLenDist * TmpVec3.y, minz + MaxLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos2.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'                TmpVec6.Set -TmpVec4.x, -TmpVec4.y, -TmpVec4.z
'                If WhichEnd = SPSMemberAxisStart Then
'                    If IsAtStart = True Then
'                        ' use tmppos2 as ref pt
'                        RefPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                        TmpVec5.Set -TmpVec3.x, -TmpVec3.y, -TmpVec3.z
'                        CopeLength = MaxLenDist
'                    Else
'                        ' use tmppos as refpt
'                        RefPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'                        TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                        CopeLength = ProjLength - MinLenDist
'                    End If
'                Else ' member end is getting cut
'                    If IsAtStart = True Then
'                        ' use tmppos as ref pt
'                        RefPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'                        TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                        CopeLength = ProjLength - MinLenDist
'                    Else
'                        ' use tmppos2 as refpt
'                        RefPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                        TmpVec5.Set -TmpVec3.x, -TmpVec3.y, -TmpVec3.z
'                        CopeLength = MaxLenDist
'                    End If
'                End If
'                If SquaredEnd = True Then
'                    DrawCornerCopeShape 1, False, Radius, RadiusType, RefPt, Nothing, TmpVec5, TmpVec6, 0#, Nothing, 0#, Nothing, curveElms
'                    DrawMode = 1
'                    PlanesConcave = False
'                    PlaneIntPt.Set RefPt.x, RefPt.y, RefPt.z
'                    Shp1Vec.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                    Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                    Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                    Clr2Vec.Set -Shp1Vec.x, -Shp1Vec.y, -Shp1Vec.z
'                    SideClr = 0#
'                    InsideClr = 0#
'                End If
'            End If
'        ElseIf Abs(MaxWidDist - ProjWidth) <= 0.0001 Then
'            If CopeType = 1 Then
'                If IsAtBottom = True Then
'                    CopeEdge = 514 ' top
'                Else
'                    CopeEdge = 513 ' bottom
'                End If
'            ElseIf CopeType = 2 And Circular = True Then
'                    If IsAtLeft = True Then
'                        CopeSide = 2 ' right
'                        CopeEdge = 514 ' top
'                    Else
'                        CopeSide = 1 ' left
'                        CopeEdge = 514 ' top
'                    End If
'            End If
'            If Abs(MinLenDist) <= 0.0001 Then
'                ' case 4 in diagram
'                If IsAtStart = True Then
'                    WhichEnd = SPSMemberAxisStart
'                Else
'                    WhichEnd = SPSMemberAxisEnd
'                End If
'                Set oCurve1 = Nothing
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MaxLenDist * TmpVec3.x, miny + MaxLenDist * TmpVec3.y, minz + MaxLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve1 = oLine
'                Set oLine = Nothing
'                TmpPos2.Set minx + MinWidDist * TmpVec4.x, miny + MinWidDist * TmpVec4.y, minz + MinWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos2.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                    TmpVec5.Set -TmpVec3.x, -TmpVec3.y, -TmpVec3.z
'                    TmpVec6.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                    If SquaredEnd = True Then
'                        DrawCornerCopeShape 1, False, Radius, RadiusType, TmpPos2, Nothing, TmpVec5, TmpVec6, 0, Nothing, 0, Nothing, curveElms
'                        DrawMode = 1
'                        PlanesConcave = False
'                        PlaneIntPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                        Shp1Vec.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                        Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                        Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                        Clr2Vec.Set -Shp1Vec.x, -Shp1Vec.y, -Shp1Vec.z
'                        SideClr = 0#
'                        InsideClr = 0#
'                    End If
'                End If
'            ElseIf Abs(MaxLenDist - ProjLength) <= 0.0001 Then
'                ' case 5 of diagram
'                If IsAtStart = True Then
'                    WhichEnd = SPSMemberAxisEnd
'                Else
'                    WhichEnd = SPSMemberAxisStart
'                End If
'                Set oCurve1 = Nothing
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MinLenDist * TmpVec3.x, miny + MinLenDist * TmpVec3.y, minz + MinLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve1 = oLine
'                Set oLine = Nothing
'                TmpPos2.Set minx + MinWidDist * TmpVec4.x, miny + MinWidDist * TmpVec4.y, minz + MinWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos2.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                    TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                    TmpVec6.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                    If SquaredEnd = True Then
'                        DrawCornerCopeShape 1, False, Radius, RadiusType, TmpPos2, Nothing, TmpVec5, TmpVec6, 0, Nothing, 0, Nothing, curveElms
'                        DrawMode = 1
'                        PlanesConcave = False
'                        PlaneIntPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                        Shp1Vec.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                        Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                        Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                        Clr2Vec.Set -Shp1Vec.x, -Shp1Vec.y, -Shp1Vec.z
'                        SideClr = 0#
'                        InsideClr = 0#
'                    End If
'                End If
'            Else
'                ' case 6 of diagram ; minlendist and maxlendist are between the extremes
'                Set oCurve1 = Nothing
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MinWidDist * TmpVec4.x, miny + MinWidDist * TmpVec4.y, minz + MinWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve1 = oLine
'                Set oLine = Nothing
'                TmpPos2.Set minx + MinLenDist * TmpVec3.x, miny + MinLenDist * TmpVec3.y, minz + MinLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MaxLenDist * TmpVec3.x, miny + MaxLenDist * TmpVec3.y, minz + MaxLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos2.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'
'                TmpVec6.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                If WhichEnd = SPSMemberAxisStart Then
'                    If IsAtStart = True Then
'                        ' use tmppos2 as ref pt
'                        RefPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                        TmpVec5.Set -TmpVec3.x, -TmpVec3.y, -TmpVec3.z
'                        CopeLength = MaxLenDist
'                    Else
'                        ' use tmppos as refpt
'                        RefPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'                        TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                        CopeLength = ProjLength - MinLenDist
'                    End If
'                Else ' member end is getting cut
'                    If IsAtStart = True Then
'                        ' use tmppos as ref pt
'                        RefPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'                        TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                        CopeLength = ProjLength - MinLenDist
'                    Else
'                        ' use tmppos2 as refpt
'                        RefPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                        TmpVec5.Set -TmpVec3.x, -TmpVec3.y, -TmpVec3.z
'                        CopeLength = MaxLenDist
'                    End If
'                End If
'                If SquaredEnd = True Then
'                    DrawCornerCopeShape 1, False, Radius, RadiusType, RefPt, Nothing, TmpVec5, TmpVec6, 0, Nothing, 0, Nothing, curveElms
'                    DrawMode = 1
'                    PlanesConcave = False
'                    PlaneIntPt.Set RefPt.x, RefPt.y, RefPt.z
'                    Shp1Vec.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                    Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                    Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                    Clr2Vec.Set -Shp1Vec.x, -Shp1Vec.y, -Shp1Vec.z
'                    SideClr = 0#
'                    InsideClr = 0#
'                End If
'            End If
'        Else ' min and max width lines are between the extremes
'            If Abs(MinLenDist) <= 0.0001 Then
'                ' case 7 in diagram
'                If IsAtStart = True Then
'                    WhichEnd = SPSMemberAxisStart
'                Else
'                    WhichEnd = SPSMemberAxisEnd
'                End If
'                Set oCurve1 = Nothing
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MaxLenDist * TmpVec3.x, miny + MaxLenDist * TmpVec3.y, minz + MaxLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve1 = oLine
'                Set oLine = Nothing
'                TmpPos2.Set minx + MinWidDist * TmpVec4.x, miny + MinWidDist * TmpVec4.y, minz + MinWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MaxWidDist * TmpVec4.x, miny + MaxWidDist * TmpVec4.y, minz + MaxWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos2.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'
'                TmpVec5.Set -TmpVec3.x, -TmpVec3.y, -TmpVec3.z
'                MinWidOther = ProjWidth - MaxWidDist
'                MaxWidOther = ProjWidth - MinWidDist
'                If MinWidDist <= MinWidOther Then
'                    ' cope towards minx, miny, minz ; use tmppos2 as refpt
'                    RefPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                    TmpVec6.Set -TmpVec4.x, -TmpVec4.y, -TmpVec4.z
'                    CopeDepth = MaxWidDist
'                    If CopeType = 1 Then 'webcope
'                        If IsAtBottom = True Then
'                            CopeEdge = 513 ' bottom
'                        Else
'                            CopeEdge = 514 ' top
'                        End If
'                    ElseIf CopeType = 2 And Circular = True Then
'                        If IsAtLeft = True Then
'                            CopeEdge = 514 ' top
'                            CopeSide = 1 ' left
'                        Else
'                            CopeEdge = 514 ' top
'                            CopeSide = 2 ' right
'                        End If
'                    End If
'                Else
'                    ' cope towards other end of the edge ; use tmppos as refpt
'                    RefPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'                    TmpVec6.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                    CopeDepth = ProjWidth - MinWidDist
'                    If CopeType = 1 Then 'webcope
'                        If IsAtBottom = True Then
'                            CopeEdge = 514 ' top
'                        Else
'                            CopeEdge = 513 ' bottom
'                        End If
'                    ElseIf CopeType = 2 And Circular = True Then
'                        If IsAtLeft = True Then
'                            CopeEdge = 514 ' top
'                            CopeSide = 2 ' right
'                        Else
'                            CopeEdge = 514 ' top
'                            CopeSide = 1 ' left
'                        End If
'                    End If
'                End If
'                If SquaredEnd = True Then
'                    DrawCornerCopeShape 1, False, Radius, RadiusType, RefPt, Nothing, TmpVec5, TmpVec6, 0, Nothing, 0, Nothing, curveElms
'                    DrawMode = 1
'                    PlanesConcave = False
'                    PlaneIntPt.Set RefPt.x, RefPt.y, RefPt.z
'                    Shp1Vec.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                    Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                    Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                    Clr2Vec.Set -Shp1Vec.x, -Shp1Vec.y, -Shp1Vec.z
'                    SideClr = 0#
'                    InsideClr = 0#
'                End If
'            ElseIf Abs(MaxLenDist - ProjLength) <= 0.0001 Then
'                ' case 8 of diagram
'                If IsAtStart = True Then
'                    WhichEnd = SPSMemberAxisEnd
'                Else
'                    WhichEnd = SPSMemberAxisStart
'                End If
'                Set oCurve1 = Nothing
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MinLenDist * TmpVec3.x, miny + MinLenDist * TmpVec3.y, minz + MinLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve1 = oLine
'                Set oLine = Nothing
'                TmpPos2.Set minx + MinWidDist * TmpVec4.x, miny + MinWidDist * TmpVec4.y, minz + MinWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MaxWidDist * TmpVec4.x, miny + MaxWidDist * TmpVec4.y, minz + MaxWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos2.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'
'                TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                MinWidOther = ProjWidth - MaxWidDist
'                MaxWidOther = ProjWidth - MinWidDist
'                If MinWidDist <= MinWidOther Then
'                    ' cope towards the minx, miny, minz side; use tmppos2 as refpt
'                    RefPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                    TmpVec6.Set -TmpVec4.x, -TmpVec4.y, -TmpVec4.z
'                    CopeDepth = MaxWidDist
'                    If CopeType = 1 Then 'webcope
'                        If IsAtBottom = True Then
'                            CopeEdge = 513 ' bottom
'                        Else
'                            CopeEdge = 514 ' top
'                        End If
'                    ElseIf CopeType = 2 And Circular = True Then
'                        If IsAtLeft = True Then
'                            CopeEdge = 514 ' top
'                            CopeSide = 1 ' left
'                        Else
'                            CopeEdge = 514 ' top
'                            CopeSide = 2 ' right
'                        End If
'                    End If
'                Else
'                    ' cope towards other end of the edge ; use tmppos as refpt
'                    RefPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'                    TmpVec6.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                    CopeDepth = ProjWidth - MinWidDist
'                    If CopeType = 1 Then 'webcope
'                        If IsAtBottom = True Then
'                            CopeEdge = 514 ' top
'                        Else
'                            CopeEdge = 513 ' bottom
'                        End If
'                    ElseIf CopeType = 2 And Circular = True Then
'                        If IsAtLeft = True Then
'                            CopeEdge = 514 ' top
'                            CopeSide = 2 ' right
'                        Else
'                            CopeEdge = 514 ' top
'                            CopeSide = 1 ' left
'                        End If
'                    End If
'                End If
'
'                If SquaredEnd = True Then
'                    DrawCornerCopeShape 1, False, Radius, RadiusType, RefPt, Nothing, TmpVec5, TmpVec6, 0, Nothing, 0, Nothing, curveElms
'                    DrawMode = 1
'                    PlanesConcave = False
'                    PlaneIntPt.Set RefPt.x, RefPt.y, RefPt.z
'                    Shp1Vec.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                    Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                    Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                    Clr2Vec.Set -Shp1Vec.x, -Shp1Vec.y, -Shp1Vec.z
'                    SideClr = 0#
'                    InsideClr = 0#
'                End If
'            Else
'                ' case 9 of diagram ; minlendist, maxlendist and minwiddist, maxwiddist are between the extremes
'                ' control should never come here as it produces a cavity in the member and not a cope
'                Set oCurve1 = Nothing
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MinWidDist * TmpVec4.x, miny + MinWidDist * TmpVec4.y, minz + MinWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve1 = oLine
'                Set oLine = Nothing
'                TmpPos2.Set minx + MinLenDist * TmpVec3.x, miny + MinLenDist * TmpVec3.y, minz + MinLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'
'                Set oCurve2 = Nothing
'                TmpPos2.Set minx + MaxLenDist * TmpVec3.x, miny + MaxLenDist * TmpVec3.y, minz + MaxLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpPos2.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'
'                Set oCurve1 = Nothing
'                TmpPos3.Set minx + MaxWidDist * TmpVec4.x, miny + MaxWidDist * TmpVec4.y, minz + MaxWidDist * TmpVec4.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos3.x, TmpPos3.y, TmpPos3.z, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'                oLine.Infinite = True
'                Set oCurve1 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpVec5.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'
'                Set oCurve2 = Nothing
'                TmpPos3.Set minx + MinLenDist * TmpVec3.x, miny + MinLenDist * TmpVec3.y, minz + MinLenDist * TmpVec3.z
'                Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, TmpPos3.x, TmpPos3.y, TmpPos3.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'                oLine.Infinite = True
'                Set oCurve2 = oLine
'                Set oLine = Nothing
'                oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
'                If NumInts <= 0 Then
'                    GoTo ErrorHandler
'                Else
'                    TmpVec6.Set IntPoints(0), IntPoints(1), IntPoints(2)
'                End If
'
'                MinWidOther = ProjWidth - MaxWidDist
'                MaxWidOther = ProjWidth - MinWidDist
'                If WhichEnd = SPSMemberAxisStart Then
'                    If IsAtStart = True Then
'                        If MinWidDist <= MinWidOther Then
'                            ' cope towards minx, miny, minz side ; use tmpvec5 as refpt
'                            RefPt.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                            CopeLength = MaxLenDist
'                            CopeDepth = MaxWidDist
'                            TmpVec5.Set -TmpVec3.x, -TmpVec3.y, -TmpVec3.z
'                            TmpVec6.Set -TmpVec4.x, -TmpVec4.y, -TmpVec4.z
'                            If CopeType = 1 Then ' web cope
'                                If IsAtBottom = True Then
'                                    CopeEdge = 513 ' bottom
'                                Else
'                                    CopeEdge = 514 ' top
'                                End If
'                            ElseIf CopeType = 2 And Circular = True Then
'                                If IsAtLeft = True Then
'                                    CopeEdge = 514 ' top
'                                    CopeSide = 1 ' left
'                                Else
'                                    CopeEdge = 514 ' top
'                                    CopeSide = 2 ' right
'                                End If
'                            End If
'                        Else
'                            ' cope towards other end of edge ; use tmppos2 as refpt
'                            RefPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                            CopeLength = MaxLenDist
'                            CopeDepth = ProjWidth - MinWidDist
'                            TmpVec5.Set -TmpVec3.x, -TmpVec3.y, -TmpVec3.z
'                            TmpVec6.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                            If CopeType = 1 Then 'web cope
'                                If IsAtBottom = True Then
'                                    CopeEdge = 514 ' top
'                                Else
'                                    CopeEdge = 513 ' bottom
'                                End If
'                            ElseIf CopeType = 2 And Circular = True Then
'                                If IsAtLeft = True Then
'                                    CopeEdge = 514 'top
'                                    CopeSide = 2 ' right
'                                Else
'                                    CopeEdge = 514 ' top
'                                    CopeSide = 1 ' left
'                                End If
'                            End If
'                        End If
'                    Else
'                        If MinWidDist <= MinWidOther Then
'                            ' cope towards minx, miny, minz side ; use tmpvec6 as refpt
'                            RefPt.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                            CopeLength = ProjLength - MinLenDist
'                            CopeDepth = MaxWidDist
'                            TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                            TmpVec6.Set -TmpVec4.x, -TmpVec4.y, -TmpVec4.z
'                            If CopeType = 1 Then ' webcope
'                                If IsAtBottom = True Then
'                                    CopeEdge = 513 ' bottom
'                                Else
'                                    CopeEdge = 514 ' top
'                                End If
'                            ElseIf CopeType = 2 And Circular = True Then
'                                If IsAtLeft = True Then
'                                    CopeEdge = 514 ' top
'                                    CopeSide = 1 ' left
'                                Else
'                                    CopeEdge = 514 ' top
'                                    CopeSide = 2 ' right
'                                End If
'                            End If
'                        Else
'                            ' cope towards other end of edge ; use tmppos as refpt
'                            RefPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'                            CopeLength = ProjLength - MinLenDist
'                            CopeDepth = ProjWidth - MinWidDist
'                            TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                            TmpVec6.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                            If CopeType = 1 Then ' web cope
'                                If IsAtBottom = True Then
'                                    CopeEdge = 514 ' top
'                                Else
'                                    CopeEdge = 513 ' bottom
'                                End If
'                            ElseIf CopeType = 2 And Circular = True Then
'                                If IsAtLeft = True Then
'                                    CopeEdge = 514 'top
'                                    CopeSide = 2 ' right
'                                Else
'                                    CopeEdge = 514 ' top
'                                    CopeSide = 1 ' left
'                                End If
'                            End If
'                        End If
'                    End If
'                Else ' member end is getting cut
'                    If IsAtStart = True Then
'                        If MinWidDist <= MinWidOther Then
'                            ' cope towards minx, miny, minz side ; use tmpvec6 as refpt
'                            RefPt.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                            CopeLength = ProjLength - MinLenDist
'                            CopeDepth = MaxWidDist
'                            TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                            TmpVec6.Set -TmpVec4.x, -TmpVec4.y, -TmpVec4.z
'                            If CopeType = 1 Then ' webcope
'                                If IsAtBottom = True Then
'                                    CopeEdge = 513 ' bottom
'                                Else
'                                    CopeEdge = 514 ' top
'                                End If
'                            ElseIf CopeType = 2 And Circular = True Then
'                                If IsAtLeft = True Then
'                                    CopeEdge = 514 ' top
'                                    CopeSide = 1 ' left
'                                Else
'                                    CopeEdge = 514 ' top
'                                    CopeSide = 2 ' right
'                                End If
'                            End If
'                        Else
'                            ' cope towards other end of edge ; use tmppos as refpt
'                            RefPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'                            CopeLength = ProjLength - MinLenDist
'                            CopeDepth = ProjWidth - MinWidDist
'                            TmpVec5.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'                            TmpVec6.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                            If CopeType = 1 Then ' webcope
'                                If IsAtBottom = True Then
'                                    CopeEdge = 514 ' top
'                                Else
'                                    CopeEdge = 513 ' bottom
'                                End If
'                            ElseIf CopeType = 1 And Circular = True Then
'                                If IsAtLeft = True Then
'                                    CopeSide = 2 ' right
'                                    CopeEdge = 514 ' top
'                                Else
'                                    CopeSide = 1 'left
'                                    CopeEdge = 514 'top
'                                End If
'                            End If
'                        End If
'                    Else ' if isatstart
'                        If MinWidDist <= MinWidOther Then
'                            ' cope towards minx, miny, minz side ; use tmpvec5 as refpt
'                            RefPt.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                            CopeLength = MaxLenDist
'                            CopeDepth = MaxWidDist
'                            TmpVec5.Set -TmpVec3.x, -TmpVec3.y, -TmpVec3.z
'                            TmpVec6.Set -TmpVec4.x, -TmpVec4.y, -TmpVec4.z
'                            If CopeType = 1 Then 'webcope
'                                If IsAtBottom = True Then
'                                    CopeEdge = 513 ' bottom
'                                Else
'                                    CopeEdge = 514 ' top
'                                End If
'                            ElseIf CopeType = 1 And Circular = True Then
'                                If IsAtLeft = True Then
'                                    CopeSide = 1 ' left
'                                    CopeEdge = 514 ' top
'                                Else
'                                    CopeSide = 2 ' right
'                                    CopeEdge = 514 ' top
'                                End If
'                            End If
'                        Else
'                            ' cope towards other end of edge ; use tmppos2 as refpt
'                            RefPt.Set TmpPos2.x, TmpPos2.y, TmpPos2.z
'                            CopeLength = MaxLenDist
'                            CopeDepth = ProjWidth - MinWidDist
'                            TmpVec5.Set -TmpVec3.x, -TmpVec3.y, TmpVec3.z
'                            TmpVec6.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                            If CopeType = 1 Then ' webcope
'                                If IsAtBottom = True Then
'                                    CopeEdge = 514 ' top
'                                Else
'                                    CopeEdge = 513 ' bottom
'                                End If
'                            ElseIf CopeType = 2 And Circular = True Then
'                                If IsAtLeft = True Then
'                                    CopeSide = 2 ' right
'                                    CopeEdge = 514 ' top
'                                Else
'                                    CopeSide = 1 'left
'                                    CopeEdge = 514 ' top
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'
'                If SquaredEnd = True Then
'                    DrawCornerCopeShape 1, False, Radius, RadiusType, RefPt, Nothing, TmpVec5, TmpVec6, 0, Nothing, 0, Nothing, curveElms
'                    DrawMode = 1
'                    PlanesConcave = False
'                    PlaneIntPt.Set RefPt.x, RefPt.y, RefPt.z
'                    Shp1Vec.Set TmpVec6.x, TmpVec6.y, TmpVec6.z
'                    Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                    Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                    Clr2Vec.Set -Shp1Vec.x, -Shp1Vec.y, -Shp1Vec.z
'                    SideClr = 0#
'                    InsideClr = 0#
'                End If
'            End If
'        End If
'    ElseIf Abs(CopeDepth - ProjWidth) <= 0.0001 And Abs(CopeLength - ProjLength) > 0.0001 Then
'        If CopeType = 1 Then
'            CopeEdge = 0 ' both top & bottom
'        ElseIf (CopeType = 2 And Circular = True) Then
'            CopeSide = 3 ' both left & right
'            CopeEdge = 514 ' top
'        End If
'        TmpDist = Sqr((PlaneIntPt.x - ProjStartPos.x) * (PlaneIntPt.x - ProjStartPos.x) + (PlaneIntPt.y - ProjStartPos.y) * (PlaneIntPt.y - ProjStartPos.y) + (PlaneIntPt.z - ProjStartPos.z) * (PlaneIntPt.z - ProjStartPos.z))
'        TmpDist2 = Sqr((PlaneIntPt.x - ProjEndPos.x) * (PlaneIntPt.x - ProjEndPos.x) + (PlaneIntPt.y - ProjEndPos.y) * (PlaneIntPt.y - ProjEndPos.y) + (PlaneIntPt.z - ProjEndPos.z) * (PlaneIntPt.z - ProjEndPos.z))
'        If TmpDist <= TmpDist2 Then
'            ' MemStart end is getting cut
'            TmpPos.Set minx + MinLenDist * TmpVec3.x, miny + MinLenDist * TmpVec3.y, minz + MinLenDist * TmpVec3.z
'            LinePos(0).Set TmpPos.x, TmpPos.y, TmpPos.z
'            LinePos(1).Set TmpPos.x + ProjWidth * TmpVec4.x, TmpPos.y + ProjWidth * TmpVec4.y, TmpPos.z + ProjWidth * TmpVec4.z
'            ProjectPtOnLine ProjStartPos, LinePos, TmpPos2
'            TmpDist = Sqr((TmpPos2.x - ProjStartPos.x) * (TmpPos2.x - ProjStartPos.x) + (TmpPos2.y - ProjStartPos.y) * (TmpPos2.y - ProjStartPos.y) + (TmpPos2.z - ProjStartPos.z) * (TmpPos2.z - ProjStartPos.z))
'
'            TmpPos.Set minx + MaxLenDist * TmpVec3.x, miny + MaxLenDist * TmpVec3.y, minz + MaxLenDist * TmpVec3.z
'            LinePos(0).Set TmpPos.x, TmpPos.y, TmpPos.z
'            LinePos(1).Set TmpPos.x + ProjWidth * TmpVec4.x, TmpPos.y + ProjWidth * TmpVec4.y, TmpPos.z + ProjWidth * TmpVec4.z
'            ProjectPtOnLine ProjStartPos, LinePos, TmpPos2
'            TmpDist2 = Sqr((TmpPos2.x - ProjStartPos.x) * (TmpPos2.x - ProjStartPos.x) + (TmpPos2.y - ProjStartPos.y) * (TmpPos2.y - ProjStartPos.y) + (TmpPos2.z - ProjStartPos.z) * (TmpPos2.z - ProjStartPos.z))
'            If TmpDist <= TmpDist2 Then
'                CopeLength = TmpDist2
'            Else
'                CopeLength = TmpDist
'            End If
'            TmpPos.Set ProjStartPos.x + CopeLength * TmpVec2.x, ProjStartPos.y + CopeLength * TmpVec2.y, ProjStartPos.z + CopeLength * TmpVec2.y
'            TmpVec5.Set -TmpVec2.x, -TmpVec2.y, -TmpVec2.z
'            If SquaredEnd = True Then
'                DrawCornerCopeShape 4, False, Radius, RadiusType, TmpPos, Nothing, TmpVec4, TmpVec5, 0, Nothing, 0, Nothing, curveElms
'                DepthCut = True
'                DrawMode = 4
'                PlanesConcave = False
'                PlaneIntPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'                Shp1Vec.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                Clr2Vec.Set 0#, 0#, 0# ' not used for drawmode 4
'                SideClr = 0#
'                InsideClr = 0#
'            End If
'        Else
'            ' MemberEnd end is getting cut
'            TmpPos.Set minx + MinLenDist * TmpVec3.x, miny + MinLenDist * TmpVec3.y, minz + MinLenDist * TmpVec3.z
'            LinePos(0).Set TmpPos.x, TmpPos.y, TmpPos.z
'            LinePos(1).Set TmpPos.x + ProjWidth * TmpVec4.x, TmpPos.y + ProjWidth * TmpVec4.y, TmpPos.z + ProjWidth * TmpVec4.z
'            ProjectPtOnLine ProjEndPos, LinePos, TmpPos2
'            TmpDist = Sqr((TmpPos2.x - ProjEndPos.x) * (TmpPos2.x - ProjEndPos.x) + (TmpPos2.y - ProjEndPos.y) * (TmpPos2.y - ProjEndPos.y) + (TmpPos2.z - ProjEndPos.z) * (TmpPos2.z - ProjEndPos.z))
'
'            TmpPos.Set minx + MaxLenDist * TmpVec3.x, miny + MaxLenDist * TmpVec3.y, minz + MaxLenDist * TmpVec3.z
'            LinePos(0).Set TmpPos.x, TmpPos.y, TmpPos.z
'            LinePos(1).Set TmpPos.x + ProjWidth * TmpVec4.x, TmpPos.y + ProjWidth * TmpVec4.y, TmpPos.z + ProjWidth * TmpVec4.z
'            ProjectPtOnLine ProjEndPos, LinePos, TmpPos2
'            TmpDist2 = Sqr((TmpPos2.x - ProjEndPos.x) * (TmpPos2.x - ProjEndPos.x) + (TmpPos2.y - ProjEndPos.y) * (TmpPos2.y - ProjEndPos.y) + (TmpPos2.z - ProjEndPos.z) * (TmpPos2.z - ProjEndPos.z))
'            If TmpDist <= TmpDist2 Then
'                CopeLength = TmpDist2
'            Else
'                CopeLength = TmpDist
'            End If
'            TmpPos.Set ProjEndPos.x + CopeLength * TmpVec2.x * -1#, ProjEndPos.y + CopeLength * TmpVec2.y * -1#, ProjEndPos.z + CopeLength * TmpVec2.z * -1#
'            TmpVec5.Set TmpVec2.x, TmpVec2.y, TmpVec2.z
'            If SquaredEnd = True Then
'                DrawCornerCopeShape 4, False, Radius, RadiusType, TmpPos, Nothing, TmpVec4, TmpVec5, 0, Nothing, 0, Nothing, curveElms
'                DepthCut = True
'                DrawMode = 4
'                PlanesConcave = False
'                PlaneIntPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'                Shp1Vec.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'                Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'                Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'                Clr2Vec.Set 0#, 0#, 0# ' not used in drawmode 4
'                SideClr = 0#
'                InsideClr = 0#
'            End If
'        End If
'    ElseIf Abs(CopeLength - ProjLength) <= 0.0001 And Abs(CopeDepth - ProjWidth) > 0.0001 Then
'        Dim LowerEndCut As Boolean
'        ' try to figure out which part has to be removed
'        If ProjIntSet(0) = True Then
'            TmpDist = Sqr((MinIntPt.x - minx) * (MinIntPt.x - minx) + (MinIntPt.y - miny) * (MinIntPt.y - miny) + (MinIntPt.z - minz) * (MinIntPt.z - minz))
'            If TmpDist <= 0.5 * ProjWidth Then
'                LowerEndCut = True
'            Else
'                LowerEndCut = False
'            End If
'        ElseIf ProjIntSet(1) = True Then
'            TmpDist = Sqr((MaxIntPt.x - maxx) * (MaxIntPt.x - maxx) + (MaxIntPt.y - maxy) * (MaxIntPt.y - maxy) + (MaxIntPt.z - maxz) * (MaxIntPt.z - maxz))
'            If TmpDist > 0.5 * ProjWidth Then
'                LowerEndCut = True
'            Else
'                LowerEndCut = False
'            End If
'        Else
'            ' determine based on planeintpt
'            LinePos(0).Set minx, miny, minz
'            LinePos(0).Set minx + ProjLength * TmpVec3.x, miny + ProjLength * TmpVec3.y, minz + ProjLength * TmpVec3.z
'            ProjectPtOnLine PlaneIntPt, LinePos, TmpPos
'            TmpDist = Sqr((TmpPos.x - PlaneIntPt.x) * (TmpPos.x - PlaneIntPt.x) + (TmpPos.y - PlaneIntPt.y) * (TmpPos.y - PlaneIntPt.y) + (TmpPos.z - PlaneIntPt.z) * (TmpPos.z - PlaneIntPt.z))
'            If TmpDist <= 0.5 * ProjWidth Then
'                LowerEndCut = True
'            Else
'                LowerEndCut = False
'            End If
'        End If
'        If LowerEndCut = True Then
'            TmpPos.Set minx + MaxWidDist * TmpVec4.x, miny + MaxWidDist * TmpVec4.y, minz + MaxWidDist * TmpVec4.z
'            TmpVec5.Set -TmpVec4.x, -TmpVec4.y, -TmpVec4.z
'            CopeDepth = MaxWidDist
'            If CopeType = 1 Then
'                If IsAtBottom = True Then
'                    CopeEdge = 513 ' bottom
'                Else
'                    CopeEdge = 514 ' top
'                End If
'            ElseIf CopeType = 2 And Circular = True Then
'                    If IsAtLeft = True Then
'                        CopeSide = 1 ' left
'                        CopeEdge = 514 ' top
'                    Else
'                        CopeSide = 2 ' right
'                        CopeEdge = 514 ' top
'                    End If
'            End If
'        Else
'            TmpPos.Set minx + MinWidDist * TmpVec4.x, miny + MinWidDist * TmpVec4.y, minz + MinWidDist * TmpVec4.z
'            TmpVec5.Set TmpVec4.x, TmpVec4.y, TmpVec4.z
'            CopeDepth = ProjWidth - MinWidDist
'            If CopeType = 1 Then
'                If IsAtBottom = True Then
'                    CopeEdge = 514 ' top
'                Else
'                    CopeEdge = 513 ' bottom
'                End If
'            ElseIf CopeType = 2 And Circular = True Then
'                    If IsAtLeft = True Then
'                        CopeSide = 2 ' right
'                        CopeEdge = 514 ' top
'                    Else
'                        CopeSide = 1 ' left
'                        CopeEdge = 514 ' top
'                    End If
'            End If
'        End If
'        px = TmpPos.x + 0.5 * ProjLength * TmpVec3.x
'        py = TmpPos.y + 0.5 * ProjLength * TmpVec3.y
'        pz = TmpPos.z + 0.5 * ProjLength * TmpVec3.z
'        TmpPos.Set px, py, pz
'        If SquaredEnd = True Then
'            DrawCornerCopeShape 4, False, Radius, RadiusType, TmpPos, Nothing, TmpVec3, TmpVec5, 0, Nothing, 0, Nothing, curveElms
'            DepthCut = False
'            DrawMode = 4
'            PlanesConcave = False
'            PlaneIntPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'            Shp1Vec.Set TmpVec3.x, TmpVec3.y, TmpVec3.z
'            Shp2Vec.Set TmpVec5.x, TmpVec5.y, TmpVec5.z
'            Clr1Vec.Set -Shp2Vec.x, -Shp2Vec.y, -Shp2Vec.z
'            Clr2Vec.Set 0#, 0#, 0# ' not used in drawmode 4
'            SideClr = 0#
'            InsideClr = 0#
'        End If
'    Else ' control should never come here
'        GoTo ErrorHandler
'    End If
'
'    ' now do the rounding of the copelength, copedepth
'    If FirstIter = True And RoundIncrement > distTol Then
'        FirstIter = False
'        ' round copeLength and copeDepth to next higher value based on increment
'        dblResult = CopeLength / RoundIncrement
'        intResult = Int(-1 * dblResult) ' round to the next higher integer ; int(-6.1) returns -7
'        NewCopeLength = (-1 * intResult) * RoundIncrement
'        dblResult = CopeDepth / RoundIncrement ' returns double value
'        intResult = Int(-1 * dblResult) ' round to the next higher integer ; int(-6.1) returns -7
'        NewCopeDepth = (-1 * intResult) * RoundIncrement
'
'        ' get the new SideClr and InsideClr values based on NewCopeLength and NewCopeDepth
'        DeltaCL = NewCopeLength - CopeLength
'        DeltaCD = NewCopeDepth - CopeDepth
'        DeltaSC = 0#
'        DeltaIC = 0#
'        If DrawMode = 1 Then
'            DeltaSC = DeltaCL * Abs(TmpVec3.Dot(Clr1Vec))
'            ICComp = DeltaSC * (Clr1Vec.Dot(Clr2Vec))
'            If ICComp < DeltaCD Then
'                DeltaIC = DeltaCD - ICComp
'            Else
'                DeltaIC = 0#
'            End If
'        ElseIf DrawMode = 4 Then
'            If DepthCut = True Then
'                DeltaSC = DeltaCD
'            Else ' lengthcut
'                DeltaSC = DeltaCL
'            End If
'        End If
'        SideClr = SideClr + DeltaSC
'        InsideClr = InsideClr + DeltaIC
'        GoTo DRAWSHAPE
'    End If
'
'    Set oCmplx2 = Nothing
'    memElms.Clear
'    Set memElms = Nothing
'
'    If CopeType = 2 Then ' flange cope
'        ComputeCopeSideEdgeForFlangeCope oSuppedPart, DrawMode, PlaneIntPt, Shp1Vec, Shp2Vec, SideClr, Clr1Vec, InsideClr, Clr2Vec, PlanesConcave, Radius, RadiusType, CopeSide, CopeEdge
'    End If
'
'    For i = 0 To 3
'        Set oStartPos(i) = Nothing
'        Set oEndPos(i) = Nothing
'    Next i
'    Set MemStartPos = Nothing
'    Set MemEndPos = Nothing
'    Set ProjStartPos = Nothing
'    Set ProjEndPos = Nothing
'    Set TmpVec1 = Nothing
'    Set TmpVec2 = Nothing
'    Set TmpVec3 = Nothing
'    Set TmpVec4 = Nothing
'    Set TmpVec5 = Nothing
'    Set TmpVec6 = Nothing
'    Set pGeometryFactory = Nothing
'    Set TmpPos = Nothing
'    Set TmpPos2 = Nothing
'    Set LinePos(0) = Nothing
'    Set LinePos(1) = Nothing
'    Set MinIntPt = Nothing
'    Set MaxIntPt = Nothing
'    Set RefPt = Nothing
'    Set TmpPos3 = Nothing
'    Exit Sub
'
'ErrorHandler:
'    HandleError MODULE, METHOD
'End Sub
Private Sub GetIntersectionPointsOfCurves(ByVal oCmplx1 As ComplexString3d, oCmplx2 As ComplexString3d, ByRef NumInts As Long, IntPoints() As Double)
    Const METHOD = "GetIntersectionPointsOfCurves"
    On Error GoTo ErrorHandler
    Dim oCurve1 As IJCurve, oCurve2 As IJCurve, i As Integer, j As Integer, k As Integer
    Dim locNumInts As Long, locIntPoints() As Double, overlap As Long, px As Double, py As Double, pz As Double
    
    NumInts = 0
    ReDim IntPoints(0) As Double

    For i = 1 To oCmplx1.CurveCount
        oCmplx1.GetCurve i, oCurve1
        For j = 1 To oCmplx2.CurveCount
            oCmplx2.GetCurve j, oCurve2
            oCurve1.Intersect oCurve2, locNumInts, locIntPoints, overlap, ISECT_UNKNOWN
            If locNumInts > 0 Then
                For k = 1 To locNumInts
                    px = locIntPoints((k - 1) * 3)
                    py = locIntPoints((k - 1) * 3 + 1)
                    pz = locIntPoints((k - 1) * 3 + 2)
                    If oCurve1.IsPointOn(px, py, pz) And oCurve2.IsPointOn(px, py, pz) Then
                        ' point is on both the curves, hence it can be treated as
                        NumInts = NumInts + 1
                        ReDim Preserve IntPoints(0 To (NumInts * 3 - 1)) As Double
                        IntPoints((NumInts - 1) * 3) = px
                        IntPoints((NumInts - 1) * 3 + 1) = py
                        IntPoints((NumInts - 1) * 3 + 2) = pz
                    End If
                Next k
            End If
            Set oCurve2 = Nothing
        Next j
        Set oCurve1 = Nothing
    Next i
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Private Sub DrawCornerCopeShape(ByVal Mode As Integer, ByVal PlanesConcave As Boolean, ByVal Radius As Double, ByVal RadiusType As Integer, ByRef RootPos1 As IJDPosition, ByVal RootPos2 As IJDPosition, ByVal Shp1Vec As IJDVector, ByVal Shp2Vec As IJDVector, ByVal SideClr As Double, ByVal Clr1Vec As IJDVector, ByVal InsideClr As Double, ByVal Clr2Vec As IJDVector, curveElms As IJElements)
    Const METHOD = "DrawCornerCopeShape"
    On Error GoTo ErrorHandler

    Dim ConsiderRadius As Boolean, idx As Integer, nextIdx As Integer
    Dim oLine3d As Line3d, oArc3d As Arc3d, oCurve1 As IJCurve, oCurve2 As IJCurve
    Dim pGeometryFactory As New GeometryFactory
    Dim oPos(1 To 13) As New DPosition, ArcCenPt As IJDPosition, ArcCenPt2 As IJDPosition, ArcCenPt3 As IJDPosition, ArcCenPt4 As IJDPosition
    Dim tmpval As Double, TmpVec1 As IJDVector, TmpVec2 As IJDVector, TmpVec3 As IJDVector, TmpVec4 As IJDVector
    Dim NumInts As Long, IntPoints() As Double, overlap As Long
    Dim dist1 As Double, dist2 As Double
    Dim px As Double, py As Double, pz As Double
    Dim oLine As IJLine
    
    Set ArcCenPt = New DPosition
    Set ArcCenPt2 = New DPosition
    Set ArcCenPt3 = New DPosition
    Set ArcCenPt4 = New DPosition
    Set TmpVec1 = New DVector
    Set TmpVec2 = New DVector
    Set TmpVec3 = New DVector
    Set TmpVec4 = New DVector
    curveElms.Clear

    If Mode = 4 Then
            oPos(3).Set RootPos1.x, RootPos1.y, RootPos1.z
            ' now apply the clearance, if any
            If Not Clr1Vec Is Nothing Then
                px = oPos(3).x + SideClr * Clr1Vec.x
                py = oPos(3).y + SideClr * Clr1Vec.y
                pz = oPos(3).z + SideClr * Clr1Vec.z
                oPos(3).Set px, py, pz
            End If
            oPos(2).Set oPos(3).x + LARGE_EDGE * Shp1Vec.x, oPos(3).y + LARGE_EDGE * Shp1Vec.y, oPos(3).z + LARGE_EDGE * Shp1Vec.z
            oPos(4).Set oPos(3).x - LARGE_EDGE * Shp1Vec.x, oPos(3).y - LARGE_EDGE * Shp1Vec.y, oPos(3).z - LARGE_EDGE * Shp1Vec.z
            oPos(1).Set oPos(2).x + (LARGE_EDGE + SideClr) * Shp2Vec.x, oPos(2).y + (LARGE_EDGE + SideClr) * Shp2Vec.y, oPos(2).z + (LARGE_EDGE + SideClr) * Shp2Vec.z
            oPos(5).Set oPos(4).x + (LARGE_EDGE + SideClr) * Shp2Vec.x, oPos(4).y + (LARGE_EDGE + SideClr) * Shp2Vec.y, oPos(4).z + (LARGE_EDGE + SideClr) * Shp2Vec.z

            Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(1).x, oPos(1).y, oPos(1).z, oPos(2).x, oPos(2).y, oPos(2).z)
            curveElms.Add oLine3d
            Set oLine3d = Nothing
            Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(2).x, oPos(2).y, oPos(2).z, oPos(4).x, oPos(4).y, oPos(4).z)
            curveElms.Add oLine3d
            Set oLine3d = Nothing
            Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(4).x, oPos(4).y, oPos(4).z, oPos(5).x, oPos(5).y, oPos(5).z)
            curveElms.Add oLine3d
            Set oLine3d = Nothing
            Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(5).x, oPos(5).y, oPos(5).z, oPos(1).x, oPos(1).y, oPos(1).z)
            curveElms.Add oLine3d
            Set oLine3d = Nothing
            Exit Sub
    ElseIf Mode = 3 Then
            oPos(1).Set RootPos1.x, RootPos1.y, RootPos1.z
            oPos(2).Set RootPos2.x, RootPos2.y, RootPos2.z
            ' the other two points are passed in shp1, shp2 directions
            oPos(3).Set Shp1Vec.x, Shp1Vec.y, Shp1Vec.z
            oPos(4).Set Shp2Vec.x, Shp2Vec.y, Shp2Vec.z
            dist1 = oPos(1).DistPt(oPos(2))
            dist2 = oPos(2).DistPt(oPos(3))
            ConsiderRadius = False
            If RadiusType <> 0 And Radius > 0.0001 Then
                If dist1 <= dist2 Then
                    If 2 * Radius < dist1 Then
                        ConsiderRadius = True
                    End If
                Else
                    If 2 * Radius < dist2 Then
                        ConsiderRadius = True
                    End If
                End If
            End If
            If ConsiderRadius = True Then
                ' get chamfer/arc parameters at point 1
                TmpVec1.Set oPos(4).x - oPos(1).x, oPos(4).y - oPos(1).y, oPos(4).z - oPos(1).z
                TmpVec1.length = 1#
                TmpVec2.Set oPos(2).x - oPos(1).x, oPos(2).y - oPos(1).y, oPos(2).z - oPos(1).z
                TmpVec2.length = 1#
                If RadiusType = 1 Then
                    GetArcPointsAndCenter oPos(1), TmpVec1, TmpVec2, Radius, oPos(6), oPos(7), ArcCenPt
                Else
                    oPos(6).Set oPos(1).x + Radius * TmpVec1.x, oPos(1).y + Radius * TmpVec1.y, oPos(1).z + Radius * TmpVec1.z
                    oPos(7).Set oPos(1).x + Radius * TmpVec2.x, oPos(1).y + Radius * TmpVec2.y, oPos(1).z + Radius * TmpVec2.z
                End If
                
                ' get chamfer/arc parameters at point 2
                TmpVec1.Set oPos(1).x - oPos(2).x, oPos(1).y - oPos(2).y, oPos(1).z - oPos(2).z
                TmpVec1.length = 1#
                TmpVec2.Set oPos(3).x - oPos(2).x, oPos(3).y - oPos(2).y, oPos(3).z - oPos(2).z
                TmpVec2.length = 1#
                If RadiusType = 1 Then
                    GetArcPointsAndCenter oPos(2), TmpVec1, TmpVec2, Radius, oPos(8), oPos(9), ArcCenPt2
                Else
                    oPos(8).Set oPos(2).x + Radius * TmpVec1.x, oPos(2).y + Radius * TmpVec1.y, oPos(2).z + Radius * TmpVec1.z
                    oPos(9).Set oPos(2).x + Radius * TmpVec2.x, oPos(2).y + Radius * TmpVec2.y, oPos(2).z + Radius * TmpVec2.z
                End If
                
                ' get chamfer/arc parameters at point 3
                TmpVec1.Set oPos(2).x - oPos(3).x, oPos(2).y - oPos(3).y, oPos(2).z - oPos(3).z
                TmpVec1.length = 1#
                TmpVec2.Set oPos(4).x - oPos(3).x, oPos(4).y - oPos(3).y, oPos(4).z - oPos(3).z
                TmpVec2.length = 1#
                If RadiusType = 1 Then
                    GetArcPointsAndCenter oPos(3), TmpVec1, TmpVec2, Radius, oPos(10), oPos(11), ArcCenPt3
                Else
                    oPos(10).Set oPos(3).x + Radius * TmpVec1.x, oPos(3).y + Radius * TmpVec1.y, oPos(3).z + Radius * TmpVec1.z
                    oPos(11).Set oPos(3).x + Radius * TmpVec2.x, oPos(3).y + Radius * TmpVec2.y, oPos(3).z + Radius * TmpVec2.z
                End If
                
                ' get arc parameters at point 4
                TmpVec1.Set oPos(3).x - oPos(4).x, oPos(3).y - oPos(4).y, oPos(3).z - oPos(4).z
                TmpVec1.length = 1#
                TmpVec2.Set oPos(1).x - oPos(4).x, oPos(1).y - oPos(4).y, oPos(1).z - oPos(4).z
                TmpVec2.length = 1#
                If RadiusType = 1 Then
                    GetArcPointsAndCenter oPos(4), TmpVec1, TmpVec2, Radius, oPos(12), oPos(13), ArcCenPt4
                Else
                    oPos(12).Set oPos(4).x + Radius * TmpVec1.x, oPos(4).y + Radius * TmpVec1.y, oPos(4).z + Radius * TmpVec1.z
                    oPos(13).Set oPos(4).x + Radius * TmpVec2.x, oPos(4).y + Radius * TmpVec2.y, oPos(4).z + Radius * TmpVec2.z
                End If

                ' now draw the shape
                Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(7).x, oPos(7).y, oPos(7).z, oPos(8).x, oPos(8).y, oPos(8).z)
                curveElms.Add oLine3d
                Set oLine3d = Nothing
                If RadiusType = 1 Then
                    Set oArc3d = pGeometryFactory.Arcs3d.CreateByCenterStartEnd(Nothing, ArcCenPt2.x, ArcCenPt2.y, ArcCenPt2.z, oPos(8).x, oPos(8).y, oPos(8).z, oPos(9).x, oPos(9).y, oPos(9).z)
                    curveElms.Add oArc3d
                    Set oArc3d = Nothing
                Else
                    Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(8).x, oPos(8).y, oPos(8).z, oPos(9).x, oPos(9).y, oPos(9).z)
                    curveElms.Add oLine3d
                    Set oLine3d = Nothing
                End If
                Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(9).x, oPos(9).y, oPos(9).z, oPos(10).x, oPos(10).y, oPos(10).z)
                curveElms.Add oLine3d
                Set oLine3d = Nothing
                If RadiusType = 1 Then
                    Set oArc3d = pGeometryFactory.Arcs3d.CreateByCenterStartEnd(Nothing, ArcCenPt3.x, ArcCenPt3.y, ArcCenPt3.z, oPos(10).x, oPos(10).y, oPos(10).z, oPos(11).x, oPos(11).y, oPos(11).z)
                    curveElms.Add oArc3d
                    Set oArc3d = Nothing
                Else
                    Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(10).x, oPos(10).y, oPos(10).z, oPos(11).x, oPos(11).y, oPos(11).z)
                    curveElms.Add oLine3d
                    Set oLine3d = Nothing
                End If
                Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(11).x, oPos(11).y, oPos(11).z, oPos(12).x, oPos(12).y, oPos(12).z)
                curveElms.Add oLine3d
                Set oLine3d = Nothing
                If RadiusType = 1 Then
                    Set oArc3d = pGeometryFactory.Arcs3d.CreateByCenterStartEnd(Nothing, ArcCenPt4.x, ArcCenPt4.y, ArcCenPt4.z, oPos(12).x, oPos(12).y, oPos(12).z, oPos(13).x, oPos(13).y, oPos(13).z)
                    curveElms.Add oArc3d
                    Set oArc3d = Nothing
                Else
                    Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(12).x, oPos(12).y, oPos(12).z, oPos(13).x, oPos(13).y, oPos(13).z)
                    curveElms.Add oLine3d
                    Set oLine3d = Nothing
                End If
                Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(13).x, oPos(13).y, oPos(13).z, oPos(6).x, oPos(6).y, oPos(6).z)
                curveElms.Add oLine3d
                Set oLine3d = Nothing
                If RadiusType = 1 Then
                    Set oArc3d = pGeometryFactory.Arcs3d.CreateByCenterStartEnd(Nothing, ArcCenPt.x, ArcCenPt.y, ArcCenPt.z, oPos(6).x, oPos(6).y, oPos(6).z, oPos(7).x, oPos(7).y, oPos(7).z)
                    curveElms.Add oArc3d
                    Set oArc3d = Nothing
                Else
                    Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(6).x, oPos(6).y, oPos(6).z, oPos(7).x, oPos(7).y, oPos(7).z)
                    curveElms.Add oLine3d
                    Set oLine3d = Nothing
                End If
            Else
                For idx = 1 To 4
                    nextIdx = (idx Mod 4) + 1
                    Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, oPos(nextIdx).x, oPos(nextIdx).y, oPos(nextIdx).z)
                    curveElms.Add oLine3d
                    Set oLine3d = Nothing
                Next idx
            End If
            Exit Sub
    End If
    ConsiderRadius = False
    If RadiusType <> 0 And Radius > 0.0001 Then
        If Mode = 2 Then
            dist1 = oPos(3).DistPt(oPos(2))
        Else
            dist1 = LARGE_EDGE
        End If
        If Mode * Radius < dist1 Then
            ConsiderRadius = True
        End If
    End If
    
    ' get shape points and make edges away from the supported member longer so that we don't get any hanging pieces
    oPos(3).Set RootPos1.x, RootPos1.y, RootPos1.z
    
    If Mode = 2 And Not RootPos2 Is Nothing Then
        oPos(2).Set RootPos2.x, RootPos2.y, RootPos2.z
    Else
        oPos(2).Set RootPos1.x + LARGE_EDGE * Shp1Vec.x, RootPos1.y + LARGE_EDGE * Shp1Vec.y, RootPos1.z + LARGE_EDGE * Shp1Vec.z
    End If

    px = RootPos1.x + LARGE_EDGE * Shp2Vec.x
    py = RootPos1.y + LARGE_EDGE * Shp2Vec.y
    pz = RootPos1.z + LARGE_EDGE * Shp2Vec.z
    oPos(4).Set px, py, pz

    If PlanesConcave = True Then
        px = oPos(2).x + 2 * LARGE_EDGE * Shp2Vec.x * -1#
        py = oPos(2).y + 2 * LARGE_EDGE * Shp2Vec.y * -1#
        pz = oPos(2).z + 2 * LARGE_EDGE * Shp2Vec.z * -1#
        oPos(1).Set px, py, pz

        px = oPos(4).x + 2 * LARGE_EDGE * Shp1Vec.x * -1#
        py = oPos(4).y + 2 * LARGE_EDGE * Shp1Vec.y * -1#
        pz = oPos(4).z + 2 * LARGE_EDGE * Shp1Vec.z * -1#
        oPos(5).Set px, py, pz
    Else
        px = oPos(2).x + LARGE_EDGE * Shp2Vec.x
        py = oPos(2).y + LARGE_EDGE * Shp2Vec.y
        pz = oPos(2).z + LARGE_EDGE * Shp2Vec.z
        oPos(1).Set px, py, pz
    End If

    ' now project the points to satisfy side/inside clearances
    If (Not Clr1Vec Is Nothing) And (Not Clr2Vec Is Nothing) Then
        TmpVec1.Set oPos(3).x - oPos(2).x, oPos(3).y - oPos(2).y, oPos(3).z - oPos(2).z
        TmpVec1.length = 1#
        TmpVec2.Set oPos(3).x - oPos(4).x, oPos(3).y - oPos(4).y, oPos(3).z - oPos(4).z
        TmpVec2.length = 1#
        
        px = oPos(2).x + SideClr * Clr1Vec.x
        py = oPos(2).y + SideClr * Clr1Vec.y
        pz = oPos(2).z + SideClr * Clr1Vec.z
        oPos(2).Set px, py, pz
    
        px = oPos(4).x + InsideClr * Clr2Vec.x
        py = oPos(4).y + InsideClr * Clr2Vec.y
        pz = oPos(4).z + InsideClr * Clr2Vec.z
        oPos(4).Set px, py, pz

        Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, oPos(4).x, oPos(4).y, oPos(4).z, TmpVec2.x, TmpVec2.y, TmpVec2.z, LARGE_EDGE)
        oLine.Infinite = True
        Set oCurve1 = oLine
        Set oLine = Nothing
        Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, oPos(2).x, oPos(2).y, oPos(2).z, TmpVec1.x, TmpVec1.y, TmpVec1.z, LARGE_EDGE)
        oLine.Infinite = True
        Set oCurve2 = oLine
        Set oLine = Nothing
        oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
        If NumInts <= 0 Then
            GoTo ErrorHandler
        End If
        oPos(3).Set IntPoints(0), IntPoints(1), IntPoints(2)
        RootPos1.Set IntPoints(0), IntPoints(1), IntPoints(2)
        Set oCurve1 = Nothing
        Set oCurve2 = Nothing
    End If

    If ConsiderRadius = True Then
        If RadiusType = 1 Then
            GetArcPointsAndCenter oPos(3), Shp1Vec, Shp2Vec, Radius, oPos(6), oPos(7), ArcCenPt
        Else
            oPos(6).Set oPos(3).x + Radius * Shp1Vec.x, oPos(3).y + Radius * Shp1Vec.y, oPos(3).z + Radius * Shp1Vec.z
            oPos(7).Set oPos(3).x + Radius * Shp2Vec.x, oPos(3).y + Radius * Shp2Vec.y, oPos(3).z + Radius * Shp2Vec.z
        End If
        If Mode = 2 Then
            TmpVec1.Set -Shp1Vec.x, -Shp1Vec.y, -Shp1Vec.z
            If RadiusType = 1 Then
                GetArcPointsAndCenter oPos(2), TmpVec1, Shp2Vec, Radius, oPos(8), oPos(9), ArcCenPt2
            Else
                oPos(8).Set oPos(2).x + Radius * TmpVec1.x, oPos(2).y + Radius * TmpVec1.y, oPos(2).z + Radius * TmpVec1.z
                oPos(9).Set oPos(2).x + Radius * Shp2Vec.x, oPos(2).y + Radius * Shp2Vec.y, oPos(2).z + Radius * Shp2Vec.z
            End If
        End If
    End If

    If ConsiderRadius = False Then
        Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(1).x, oPos(1).y, oPos(1).z, oPos(2).x, oPos(2).y, oPos(2).z)
        curveElms.Add oLine3d
        Set oLine3d = Nothing
        Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(2).x, oPos(2).y, oPos(2).z, oPos(3).x, oPos(3).y, oPos(3).z)
        curveElms.Add oLine3d
        Set oLine3d = Nothing
        Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(3).x, oPos(3).y, oPos(3).z, oPos(4).x, oPos(4).y, oPos(4).z)
        curveElms.Add oLine3d
        Set oLine3d = Nothing
    Else
        If Mode = 2 Then
            Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(1).x, oPos(1).y, oPos(1).z, oPos(9).x, oPos(9).y, oPos(9).z)
            curveElms.Add oLine3d
            Set oLine3d = Nothing
            If RadiusType = 1 Then
                Set oArc3d = pGeometryFactory.Arcs3d.CreateByCenterStartEnd(Nothing, ArcCenPt2.x, ArcCenPt2.y, ArcCenPt2.z, oPos(9).x, oPos(9).y, oPos(9).z, oPos(8).x, oPos(8).y, oPos(8).z)
                curveElms.Add oArc3d
                Set oArc3d = Nothing
            Else
                Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(9).x, oPos(9).y, oPos(9).z, oPos(8).x, oPos(8).y, oPos(8).z)
                curveElms.Add oLine3d
                Set oLine3d = Nothing
            End If
            Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(8).x, oPos(8).y, oPos(8).z, oPos(6).x, oPos(6).y, oPos(6).z)
            curveElms.Add oLine3d
            Set oLine3d = Nothing
        Else
            Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(1).x, oPos(1).y, oPos(1).z, oPos(2).x, oPos(2).y, oPos(2).z)
            curveElms.Add oLine3d
            Set oLine3d = Nothing
            Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(2).x, oPos(2).y, oPos(2).z, oPos(6).x, oPos(6).y, oPos(6).z)
            curveElms.Add oLine3d
            Set oLine3d = Nothing
        End If
        If RadiusType = 1 Then
            Set oArc3d = pGeometryFactory.Arcs3d.CreateByCenterStartEnd(Nothing, ArcCenPt.x, ArcCenPt.y, ArcCenPt.z, oPos(6).x, oPos(6).y, oPos(6).z, oPos(7).x, oPos(7).y, oPos(7).z)
            curveElms.Add oArc3d
            Set oArc3d = Nothing
        Else
            Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(6).x, oPos(6).y, oPos(6).z, oPos(7).x, oPos(7).y, oPos(7).z)
            curveElms.Add oLine3d
            Set oLine3d = Nothing
        End If
        Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(7).x, oPos(7).y, oPos(7).z, oPos(4).x, oPos(4).y, oPos(4).z)
        curveElms.Add oLine3d
        Set oLine3d = Nothing
    End If
    
    If PlanesConcave = False Then
        Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(4).x, oPos(4).y, oPos(4).z, oPos(1).x, oPos(1).y, oPos(1).z)
        curveElms.Add oLine3d
        Set oLine3d = Nothing
    Else
        Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(4).x, oPos(4).y, oPos(4).z, oPos(5).x, oPos(5).y, oPos(5).z)
        curveElms.Add oLine3d
        Set oLine3d = Nothing
        Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(5).x, oPos(5).y, oPos(5).z, oPos(1).x, oPos(1).y, oPos(1).z)
        curveElms.Add oLine3d
        Set oLine3d = Nothing
    End If

    For idx = 1 To 13
        Set oPos(idx) = Nothing
    Next idx
    Set ArcCenPt = Nothing
    Set ArcCenPt2 = Nothing
    Set ArcCenPt3 = Nothing
    Set ArcCenPt4 = Nothing
    Set TmpVec1 = Nothing
    Set TmpVec2 = Nothing
    Set TmpVec3 = Nothing
    Set TmpVec4 = Nothing
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub ComputeSurfaceTrimShapeAndOutputs(oSuppedPart As ISPSMemberPartPrismatic, ByVal SquaredEnd As Boolean, WhichEnd As SPSMemberAxisPortIndex, pResourceManager As IUnknown, ByRef oCutbackSurf As Object, ByRef FlangeAngle As Double, ByRef WebAngle As Double) ' TR#79729 ; changed from IJSurface
    Const METHOD = "ComputeSurfaceTrimShapeAndOutputs"
    On Error GoTo ErrorHandler
    Dim oTopFacePos(1 To 4) As IJDPosition, oBotFacePos(1 To 4) As IJDPosition
    Dim i As Integer, nextIdx As Integer, oCurve As IJCurve, oLine As IJLine
    Dim pGeometryFactory As New GeometryFactory
    Dim MinLenTop As Double, MaxLenTop As Double, MinWidTop As Double, MaxWidTop As Double
    Dim MinLenBot As Double, MaxLenBot As Double, MinWidBot As Double, MaxWidBot As Double
    Dim MinLenFront As Double, MaxLenFront As Double, MinWidFront As Double, MaxWidFront As Double
    Dim MinLenBack As Double, MaxLenBack As Double, MinWidBack As Double, MaxWidBack As Double
    Dim MinWidLenTop As Double, MaxWidLenTop As Double, MinLenWidTop As Double, MaxLenWidTop As Double
    Dim MinWidLenBot As Double, MaxWidLenBot As Double, MinLenWidBot As Double, MaxLenWidBot As Double
    Dim MinWidLenFront As Double, MaxWidLenFront As Double, MinLenWidFront As Double, MaxLenWidFront As Double
    Dim MinWidLenBack As Double, MaxWidLenBack As Double, MinLenWidBack As Double, MaxLenWidBack As Double
    Dim intElms As IJElements, code As Geom3dIntersectConstants
    Dim ProjLength As Double, ProjWidth As Double, StartPos As IJDPosition, EndPos As IJDPosition
    Dim px As Double, py As Double, pz As Double, IntPoint As IJPoint, TmpPos1 As IJDPosition, TmpPos2 As IJDPosition, TmpPos3 As IJDPosition
    Dim Line1Pos(0 To 1) As IJDPosition, Line2Pos(0 To 1) As IJDPosition
    Dim TmpDist1 As Double, TmpDist2 As Double
    Dim oPlaneRoot As IJDPosition, oPlaneNormal As IJDVector, MemVec As IJDVector, TmpVec As IJDVector, TmpVec1 As IJDVector, TmpVec2 As IJDVector
    Dim CutLength As Double, CutDepth As Double, CosAngle As Double, ProjHeight As Double
    Dim IntersectsTop As Boolean, IntersectsBot As Boolean, IntersectsFront As Boolean, IntersectsBack As Boolean
    Dim BotMinLenFlag As Boolean, BotMaxLenFlag As Boolean, BotWidFlag As Boolean
    Dim myobj As Object, IsPlanar As Boolean
    Dim j As Integer ' TR#75014
    Dim iSurf As IJSurface ' TR#79729
    
    Set StartPos = New DPosition
    Set EndPos = New DPosition
    Set TmpPos1 = New DPosition
    Set TmpPos2 = New DPosition
    Set TmpPos3 = New DPosition
    Set oPlaneRoot = New DPosition
    Set oPlaneNormal = New DVector
    Set MemVec = New DVector
    Set TmpVec = New DVector
    Set TmpVec1 = New DVector
    Set TmpVec2 = New DVector
    Set Line1Pos(0) = New DPosition
    Set Line1Pos(1) = New DPosition
    Set Line2Pos(0) = New DPosition
    Set Line2Pos(1) = New DPosition
    Set oTopFacePos(1) = New DPosition
    Set oTopFacePos(2) = New DPosition
    Set oTopFacePos(3) = New DPosition
    Set oTopFacePos(4) = New DPosition
    Set oBotFacePos(1) = New DPosition
    Set oBotFacePos(2) = New DPosition
    Set oBotFacePos(3) = New DPosition
    Set oBotFacePos(4) = New DPosition
    
    Set myobj = oCutbackSurf
    IsPlanar = True
    If Not myobj Is Nothing Then
        If Not (TypeOf myobj Is IJPlane) Then
            IsPlanar = False
        End If
    End If
    Set myobj = Nothing
    Set iSurf = oCutbackSurf ' TR#79729
    
    GetFacePoints oSuppedPart, 1, False, oTopFacePos
    GetFacePoints oSuppedPart, 2, False, oBotFacePos

    Set oLine = oSuppedPart
    oLine.GetStartPoint px, py, pz
    StartPos.Set px, py, pz
    oLine.GetEndPoint px, py, pz
    EndPos.Set px, py, pz
    Set oLine = Nothing
    MemVec.Set EndPos.x - StartPos.x, EndPos.y - StartPos.y, EndPos.z - StartPos.z
    MemVec.length = 1#
    ProjLength = StartPos.DistPt(EndPos)
    ProjWidth = oTopFacePos(1).DistPt(oTopFacePos(2))
    ProjHeight = oTopFacePos(1).DistPt(oBotFacePos(1))

    ' the following logic may have to be modified in case of tapered members which have
    ' unequal section sizes. The current logic works only for members with uniform cross-section.
    IntersectsTop = False
    IntersectsBot = False
    IntersectsFront = False
    IntersectsBack = False
    
    ' get the int points max/min at the top face
    MinLenTop = ProjLength
    MaxLenTop = 0
    MinWidTop = ProjWidth
    MaxWidTop = 0
    Line1Pos(0).Set oTopFacePos(1).x, oTopFacePos(1).y, oTopFacePos(1).z
    Line1Pos(1).Set oTopFacePos(2).x, oTopFacePos(2).y, oTopFacePos(2).z
    Line2Pos(0).Set oTopFacePos(1).x, oTopFacePos(1).y, oTopFacePos(1).z
    Line2Pos(1).Set oTopFacePos(4).x, oTopFacePos(4).y, oTopFacePos(4).z
    For i = 1 To 4
        nextIdx = (i Mod 4) + 1
        Set oCurve = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oTopFacePos(i).x, oTopFacePos(i).y, oTopFacePos(i).z, oTopFacePos(nextIdx).x, oTopFacePos(nextIdx).y, oTopFacePos(nextIdx).z)
        iSurf.Intersect oCurve, intElms, code ' TR#79729
        If Not intElms Is Nothing Then
            If intElms.count > 0 Then
                For j = 1 To intElms.count ' TR#75014 ; loop thru all point especially in case of circular
                    Set IntPoint = intElms.Item(j)
                    IntPoint.GetPoint px, py, pz
                    If oCurve.IsPointOn(px, py, pz) Then
                        TmpPos1.Set px, py, pz
                        ProjectPtOnLine TmpPos1, Line1Pos, TmpPos2
                        TmpDist1 = Line1Pos(0).DistPt(TmpPos2)
                        ProjectPtOnLine TmpPos1, Line2Pos, TmpPos3
                        TmpDist2 = Line2Pos(0).DistPt(TmpPos3)
                        
                        If MinWidTop > TmpDist1 Then
                            MinWidTop = TmpDist1
                            MinWidLenTop = TmpDist2
                        End If
                        If MaxWidTop < TmpDist1 Then
                            MaxWidTop = TmpDist1
                            MaxWidLenTop = TmpDist2
                        End If
                        If MinLenTop > TmpDist2 Then
                            MinLenTop = TmpDist2
                            MinLenWidTop = TmpDist1
                        End If
                        If MaxLenTop < TmpDist2 Then
                            MaxLenTop = TmpDist2
                            MaxLenWidTop = TmpDist1
                        End If
                        IntersectsTop = True
                    End If
                Next j
            End If
        End If
        Set oCurve = Nothing
    Next i
    If Abs(MaxLenTop - MinLenTop) <= 0.0001 Then
        CutLength = MaxLenTop
    Else
        CutLength = MaxLenTop - MinLenTop
    End If
    If Abs(MaxWidTop - MinWidTop) <= 0.0001 Then
        CutDepth = MaxWidTop
    Else
        CutDepth = MaxWidTop - MinWidTop
    End If

    ' get the int points max/min at the bottom face
    MinLenBot = ProjLength
    MaxLenBot = 0
    MinWidBot = ProjWidth
    MaxWidBot = 0
    Line1Pos(0).Set oBotFacePos(1).x, oBotFacePos(1).y, oBotFacePos(1).z
    Line1Pos(1).Set oBotFacePos(2).x, oBotFacePos(2).y, oBotFacePos(2).z
    
    Line2Pos(0).Set oBotFacePos(1).x, oBotFacePos(1).y, oBotFacePos(1).z
    Line2Pos(1).Set oBotFacePos(4).x, oBotFacePos(4).y, oBotFacePos(4).z
    For i = 1 To 4
        nextIdx = (i Mod 4) + 1
        Set oCurve = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oBotFacePos(i).x, oBotFacePos(i).y, oBotFacePos(i).z, oBotFacePos(nextIdx).x, oBotFacePos(nextIdx).y, oBotFacePos(nextIdx).z)
        iSurf.Intersect oCurve, intElms, code ' TR#79729
        If Not intElms Is Nothing Then
            If intElms.count > 0 Then
                For j = 1 To intElms.count ' TR#75014 ; loop thru all point especially in case of circular
                    Set IntPoint = intElms.Item(j)
                    IntPoint.GetPoint px, py, pz
                    If oCurve.IsPointOn(px, py, pz) Then
                        TmpPos1.Set px, py, pz
                        ProjectPtOnLine TmpPos1, Line1Pos, TmpPos2
                        TmpDist1 = Line1Pos(0).DistPt(TmpPos2)
                        ProjectPtOnLine TmpPos1, Line2Pos, TmpPos3
                        TmpDist2 = Line2Pos(0).DistPt(TmpPos3)
                        
                        If MinWidBot > TmpDist1 Then
                            MinWidBot = TmpDist1
                            MinWidLenBot = TmpDist2
                        End If
                        If MaxWidBot < TmpDist1 Then
                            MaxWidBot = TmpDist1
                            MaxWidLenBot = TmpDist2
                        End If
                        If MinLenBot > TmpDist2 Then
                            MinLenBot = TmpDist2
                            MinLenWidBot = TmpDist1
                        End If
                        If MaxLenBot < TmpDist2 Then
                            MaxLenBot = TmpDist2
                            MaxLenWidBot = TmpDist1
                        End If
                        IntersectsBot = True
                    End If
                Next j
            End If
        End If
        Set oCurve = Nothing
    Next i
    If Abs(MaxLenBot - MinLenBot) <= 0.0001 Then
        TmpDist1 = MaxLenBot
    Else
        TmpDist1 = MaxLenBot - MinLenBot
    End If
    BotMinLenFlag = False
    BotMaxLenFlag = False
    If CutLength < TmpDist1 Then
        CutLength = TmpDist1
    End If
    If MinLenBot < MinLenTop Then
        BotMinLenFlag = True
    End If
    If MaxLenBot > MaxLenTop Then
        BotMaxLenFlag = True
    End If
    If Abs(MaxWidBot - MinWidBot) <= 0.0001 Then
        TmpDist2 = MaxWidBot
    Else
        TmpDist2 = MaxWidBot - MinWidBot
    End If
    BotWidFlag = False
    If CutDepth < TmpDist2 Then
        CutDepth = TmpDist2
        BotWidFlag = True
    End If

    ' get the int points max/min at the front face
    MinLenFront = ProjLength
    MaxLenFront = 0
    MinWidFront = ProjHeight
    MaxWidFront = 0
    Line1Pos(0).Set oTopFacePos(2).x, oTopFacePos(2).y, oTopFacePos(2).z
    Line1Pos(1).Set oBotFacePos(2).x, oBotFacePos(2).y, oBotFacePos(2).z
    Line2Pos(0).Set oTopFacePos(2).x, oTopFacePos(2).y, oTopFacePos(2).z
    Line2Pos(1).Set oTopFacePos(3).x, oTopFacePos(3).y, oTopFacePos(3).z
    For i = 1 To 4
        Select Case i
            Case 1
                TmpPos2.Set oTopFacePos(2).x, oTopFacePos(2).y, oTopFacePos(2).z
                TmpPos3.Set oBotFacePos(2).x, oBotFacePos(2).y, oBotFacePos(2).z
            Case 2
                TmpPos2.Set oBotFacePos(2).x, oBotFacePos(2).y, oBotFacePos(2).z
                TmpPos3.Set oBotFacePos(3).x, oBotFacePos(3).y, oBotFacePos(3).z
            Case 3
                TmpPos2.Set oBotFacePos(3).x, oBotFacePos(3).y, oBotFacePos(3).z
                TmpPos3.Set oTopFacePos(3).x, oTopFacePos(3).y, oTopFacePos(3).z
            Case 4
                TmpPos2.Set oTopFacePos(3).x, oTopFacePos(3).y, oTopFacePos(3).z
                TmpPos3.Set oTopFacePos(2).x, oTopFacePos(2).y, oTopFacePos(2).z
        End Select
        Set oCurve = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpPos3.x, TmpPos3.y, TmpPos3.z)
        iSurf.Intersect oCurve, intElms, code ' TR#79729
        If Not intElms Is Nothing Then
            If intElms.count > 0 Then
                For j = 1 To intElms.count ' TR#75014 ; loop thru all point especially in case of circular
                    Set IntPoint = intElms.Item(j)
                    IntPoint.GetPoint px, py, pz
                    If oCurve.IsPointOn(px, py, pz) Then
                        TmpPos1.Set px, py, pz
                        ProjectPtOnLine TmpPos1, Line1Pos, TmpPos2
                        TmpDist1 = Line1Pos(0).DistPt(TmpPos2)
                        ProjectPtOnLine TmpPos1, Line2Pos, TmpPos3
                        TmpDist2 = Line2Pos(0).DistPt(TmpPos3)
                        
                        If MinWidFront > TmpDist1 Then
                            MinWidFront = TmpDist1
                            MinWidLenFront = TmpDist2
                        End If
                        If MaxWidFront < TmpDist1 Then
                            MaxWidFront = TmpDist1
                            MaxWidLenFront = TmpDist2
                        End If
                        If MinLenFront > TmpDist2 Then
                            MinLenFront = TmpDist2
                            MinLenWidFront = TmpDist1
                        End If
                        If MaxLenFront < TmpDist2 Then
                            MaxLenFront = TmpDist2
                            MaxLenWidFront = TmpDist1
                        End If
                        IntersectsFront = True
                    End If
                Next j
            End If
        End If
        Set oCurve = Nothing
    Next i

    ' get the int points max/min at the back face
    MinLenBack = ProjLength
    MaxLenBack = 0
    MinWidBack = ProjHeight
    MaxWidBack = 0
    Line1Pos(0).Set oTopFacePos(1).x, oTopFacePos(1).y, oTopFacePos(1).z
    Line1Pos(1).Set oBotFacePos(1).x, oBotFacePos(1).y, oBotFacePos(1).z
    Line2Pos(0).Set oTopFacePos(1).x, oTopFacePos(1).y, oTopFacePos(1).z
    Line2Pos(1).Set oTopFacePos(4).x, oTopFacePos(4).y, oTopFacePos(4).z
    For i = 1 To 4
        Select Case i
            Case 1
                TmpPos2.Set oTopFacePos(1).x, oTopFacePos(1).y, oTopFacePos(1).z
                TmpPos3.Set oBotFacePos(1).x, oBotFacePos(1).y, oBotFacePos(1).z
            Case 2
                TmpPos2.Set oBotFacePos(1).x, oBotFacePos(1).y, oBotFacePos(1).z
                TmpPos3.Set oBotFacePos(4).x, oBotFacePos(4).y, oBotFacePos(4).z
            Case 3
                TmpPos2.Set oBotFacePos(4).x, oBotFacePos(4).y, oBotFacePos(4).z
                TmpPos3.Set oTopFacePos(4).x, oTopFacePos(4).y, oTopFacePos(4).z
            Case 4
                TmpPos2.Set oTopFacePos(4).x, oTopFacePos(4).y, oTopFacePos(4).z
                TmpPos3.Set oTopFacePos(1).x, oTopFacePos(1).y, oTopFacePos(1).z
        End Select
        Set oCurve = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, TmpPos2.x, TmpPos2.y, TmpPos2.z, TmpPos3.x, TmpPos3.y, TmpPos3.z)
        iSurf.Intersect oCurve, intElms, code ' TR#79729
        If Not intElms Is Nothing Then
            If intElms.count > 0 Then
                For j = 1 To intElms.count ' TR#75014 ; loop thru all point especially in case of circular
                    Set IntPoint = intElms.Item(j)
                    IntPoint.GetPoint px, py, pz
                    If oCurve.IsPointOn(px, py, pz) Then
                        TmpPos1.Set px, py, pz
                        ProjectPtOnLine TmpPos1, Line1Pos, TmpPos2
                        TmpDist1 = Line1Pos(0).DistPt(TmpPos2)
                        ProjectPtOnLine TmpPos1, Line2Pos, TmpPos3
                        TmpDist2 = Line2Pos(0).DistPt(TmpPos3)
                        
                        If MinWidBack > TmpDist1 Then
                            MinWidBack = TmpDist1
                            MinWidLenBack = TmpDist2
                        End If
                        If MaxWidBack < TmpDist1 Then
                            MaxWidBack = TmpDist1
                            MaxWidLenBack = TmpDist2
                        End If
                        If MinLenBack > TmpDist2 Then
                            MinLenBack = TmpDist2
                            MinLenWidBack = TmpDist1
                        End If
                        If MaxLenBack < TmpDist2 Then
                            MaxLenBack = TmpDist2
                            MaxLenWidBack = TmpDist1
                        End If
                        IntersectsBack = True
                    End If
                Next j
            End If
        End If
        Set oCurve = Nothing
    Next i
    Set iSurf = Nothing ' TR#79729

    ' now compute the flange and web angles
    If SquaredEnd = True Or IsPlanar = False Then
        FlangeAngle = 0#
        WebAngle = 0#
    Else ' compute based on existing plane
        If IntersectsTop = True Then
            Set TmpVec1 = oTopFacePos(2).Subtract(oTopFacePos(1))
            TmpVec1.length = 1#
            Set TmpVec2 = oTopFacePos(4).Subtract(oTopFacePos(1))
            TmpVec2.length = 1#
            If Abs(MaxLenTop - MinLenTop - ProjLength) <= 0.0001 Then
                px = oTopFacePos(1).x + MinLenTop * TmpVec2.x
                py = oTopFacePos(1).y + MinLenTop * TmpVec2.y
                pz = oTopFacePos(1).z + MinLenTop * TmpVec2.z
                TmpPos1.Set px + MinLenWidTop * TmpVec1.x, py + MinLenWidTop * TmpVec1.y, pz + MinLenWidTop * TmpVec1.z
                px = oTopFacePos(1).x + MaxLenTop * TmpVec2.x
                py = oTopFacePos(1).y + MaxLenTop * TmpVec2.y
                pz = oTopFacePos(1).z + MaxLenTop * TmpVec2.z
                TmpPos2.Set px + MaxWidLenTop * TmpVec2.x, py + MaxWidLenTop * TmpVec2.y, pz + MaxWidLenTop * TmpVec2.z
            Else
                px = oTopFacePos(1).x + MinWidTop * TmpVec1.x
                py = oTopFacePos(1).y + MinWidTop * TmpVec1.y
                pz = oTopFacePos(1).z + MinWidTop * TmpVec1.z
                TmpPos1.Set px + MinWidLenTop * TmpVec2.x, py + MinWidLenTop * TmpVec2.y, pz + MinWidLenTop * TmpVec2.z
                px = oTopFacePos(1).x + MaxWidTop * TmpVec1.x
                py = oTopFacePos(1).y + MaxWidTop * TmpVec1.y
                pz = oTopFacePos(1).z + MaxWidTop * TmpVec1.z
                TmpPos2.Set px + MaxWidLenTop * TmpVec2.x, py + MaxWidLenTop * TmpVec2.y, pz + MaxWidLenTop * TmpVec2.z
            End If
            Set TmpVec2 = TmpPos1.Subtract(TmpPos2)
            TmpVec2.length = 1#
            If (Abs(TmpVec1.x - TmpVec2.x) <= 0.0001 And Abs(TmpVec1.y - TmpVec2.y) <= 0.0001 And Abs(TmpVec1.z - TmpVec2.z) <= 0.0001) Or _
                (Abs(TmpVec1.x + TmpVec2.x) <= 0.0001 And Abs(TmpVec1.y + TmpVec2.y) <= 0.0001 And Abs(TmpVec1.z + TmpVec2.z) <= 0.0001) Then
                CosAngle = 1#
            Else
                CosAngle = Abs(TmpVec1.x * TmpVec2.x + TmpVec1.y * TmpVec2.y + TmpVec1.z * TmpVec2.z)
            End If
            If Abs(CosAngle) <= 0.0001 Then
                FlangeAngle = 90#
            Else
                FlangeAngle = 180# * (Atn(Sqr(1 - CosAngle * CosAngle) / CosAngle)) / PI
            End If
        Else ' intersects bottom
            Set TmpVec1 = oBotFacePos(2).Subtract(oBotFacePos(1))
            TmpVec1.length = 1#
            Set TmpVec2 = oBotFacePos(4).Subtract(oBotFacePos(1))
            TmpVec2.length = 1#
            If Abs(MaxLenBot - MinLenBot - ProjLength) <= 0.0001 Then
                px = oBotFacePos(1).x + MinLenBot * TmpVec2.x
                py = oBotFacePos(1).y + MinLenBot * TmpVec2.y
                pz = oBotFacePos(1).z + MinLenBot * TmpVec2.z
                TmpPos1.Set px + MinLenWidBot * TmpVec1.x, py + MinLenWidBot * TmpVec1.y, pz + MinLenWidBot * TmpVec1.z
                px = oBotFacePos(1).x + MaxLenBot * TmpVec2.x
                py = oBotFacePos(1).y + MaxLenBot * TmpVec2.y
                pz = oBotFacePos(1).z + MaxLenBot * TmpVec2.z
                TmpPos2.Set px + MaxWidLenBot * TmpVec2.x, py + MaxWidLenBot * TmpVec2.y, pz + MaxWidLenBot * TmpVec2.z
            Else
                px = oBotFacePos(1).x + MinWidBot * TmpVec1.x
                py = oBotFacePos(1).y + MinWidBot * TmpVec1.y
                pz = oBotFacePos(1).z + MinWidBot * TmpVec1.z
                TmpPos1.Set px + MinWidLenBot * TmpVec2.x, py + MinWidLenBot * TmpVec2.y, pz + MinWidLenBot * TmpVec2.z
                px = oBotFacePos(1).x + MaxWidTop * TmpVec1.x
                py = oBotFacePos(1).y + MaxWidTop * TmpVec1.y
                pz = oBotFacePos(1).z + MaxWidTop * TmpVec1.z
                TmpPos2.Set px + MaxWidLenBot * TmpVec2.x, py + MaxWidLenBot * TmpVec2.y, pz + MaxWidLenBot * TmpVec2.z
            End If
            Set TmpVec2 = TmpPos1.Subtract(TmpPos2)
            TmpVec2.length = 1#
            If (Abs(TmpVec1.x - TmpVec2.x) <= 0.0001 And Abs(TmpVec1.y - TmpVec2.y) <= 0.0001 And Abs(TmpVec1.z - TmpVec2.z) <= 0.0001) Or _
                (Abs(TmpVec1.x + TmpVec2.x) <= 0.0001 And Abs(TmpVec1.y + TmpVec2.y) <= 0.0001 And Abs(TmpVec1.z + TmpVec2.z) <= 0.0001) Then
                CosAngle = 1#
            Else
                CosAngle = Abs(TmpVec1.x * TmpVec2.x + TmpVec1.y * TmpVec2.y + TmpVec1.z * TmpVec2.z)
            End If
            If Abs(CosAngle) <= 0.0001 Then
                FlangeAngle = 90#
            Else
                FlangeAngle = 180# * (Atn(Sqr(1 - CosAngle * CosAngle) / CosAngle)) / PI
            End If
        End If
        
        If IntersectsFront = True Then
            Set TmpVec1 = oBotFacePos(2).Subtract(oTopFacePos(2))
            TmpVec1.length = 1#
            Set TmpVec2 = oTopFacePos(3).Subtract(oTopFacePos(2))
            TmpVec2.length = 1#
            
            If Abs(MaxLenFront - MinLenFront - ProjLength) <= 0.0001 Then
                px = oTopFacePos(2).x + MinLenFront * TmpVec2.x
                py = oTopFacePos(2).y + MinLenFront * TmpVec2.y
                pz = oTopFacePos(2).z + MinLenFront * TmpVec2.z
                TmpPos1.Set px + MinLenWidFront * TmpVec1.x, py + MinLenWidFront * TmpVec1.y, pz + MinLenWidFront * TmpVec1.z
                px = oTopFacePos(2).x + MaxLenFront * TmpVec2.x
                py = oTopFacePos(2).y + MaxLenFront * TmpVec2.y
                pz = oTopFacePos(2).z + MaxLenFront * TmpVec2.z
                TmpPos2.Set px + MaxLenWidFront * TmpVec1.x, py + MaxLenWidFront * TmpVec1.y, pz + MaxLenWidFront * TmpVec1.z
            Else
                px = oTopFacePos(2).x + MinWidFront * TmpVec1.x
                py = oTopFacePos(2).y + MinWidFront * TmpVec1.y
                pz = oTopFacePos(2).z + MinWidFront * TmpVec1.z
                TmpPos1.Set px + MinWidLenFront * TmpVec2.x, py + MinWidLenFront * TmpVec2.y, pz + MinWidLenFront * TmpVec2.z
                px = oTopFacePos(2).x + MaxWidFront * TmpVec1.x
                py = oTopFacePos(2).y + MaxWidFront * TmpVec1.y
                pz = oTopFacePos(2).z + MaxWidFront * TmpVec1.z
                TmpPos2.Set px + MaxWidLenFront * TmpVec2.x, py + MaxWidLenFront * TmpVec2.y, pz + MaxWidLenFront * TmpVec2.z
            End If
            Set TmpVec2 = TmpPos1.Subtract(TmpPos2)
            TmpVec2.length = 1#
            If (Abs(TmpVec1.x - TmpVec2.x) <= 0.0001 And Abs(TmpVec1.y - TmpVec2.y) <= 0.0001 And Abs(TmpVec1.z - TmpVec2.z) <= 0.0001) Or _
                (Abs(TmpVec1.x + TmpVec2.x) <= 0.0001 And Abs(TmpVec1.y + TmpVec2.y) <= 0.0001 And Abs(TmpVec1.z + TmpVec2.z) <= 0.0001) Then
                CosAngle = 1#
            Else
                CosAngle = Abs(TmpVec1.x * TmpVec2.x + TmpVec1.y * TmpVec2.y + TmpVec1.z * TmpVec2.z)
            End If
            If Abs(CosAngle) <= 0.0001 Then
                WebAngle = 90#
            Else
                WebAngle = 180# * (Atn(Sqr(1 - CosAngle * CosAngle) / CosAngle)) / PI
            End If
        Else ' intersects back
            Set TmpVec1 = oBotFacePos(1).Subtract(oTopFacePos(1))
            TmpVec1.length = 1#
            Set TmpVec2 = oTopFacePos(4).Subtract(oTopFacePos(1))
            TmpVec2.length = 1#
            
            If Abs(MaxLenBack - MinLenBack - ProjLength) <= 0.0001 Then
                px = oTopFacePos(1).x + MinLenBack * TmpVec2.x
                py = oTopFacePos(1).y + MinLenBack * TmpVec2.y
                pz = oTopFacePos(1).z + MinLenBack * TmpVec2.z
                TmpPos1.Set px + MinLenWidBack * TmpVec1.x, py + MinLenWidBack * TmpVec1.y, pz + MinLenWidBack * TmpVec1.z
                px = oTopFacePos(1).x + MaxLenBack * TmpVec2.x
                py = oTopFacePos(1).y + MaxLenBack * TmpVec2.y
                pz = oTopFacePos(1).z + MaxLenBack * TmpVec2.z
                TmpPos2.Set px + MaxLenWidBack * TmpVec1.x, py + MaxLenWidBack * TmpVec1.y, pz + MaxLenWidBack * TmpVec1.z
            Else
                px = oTopFacePos(1).x + MinWidBack * TmpVec1.x
                py = oTopFacePos(1).y + MinWidBack * TmpVec1.y
                pz = oTopFacePos(1).z + MinWidBack * TmpVec1.z
                TmpPos1.Set px + MinWidLenBack * TmpVec2.x, py + MinWidLenBack * TmpVec2.y, pz + MinWidLenBack * TmpVec2.z
                px = oTopFacePos(1).x + MaxWidBack * TmpVec1.x
                py = oTopFacePos(1).y + MaxWidBack * TmpVec1.y
                pz = oTopFacePos(1).z + MaxWidBack * TmpVec1.z
                TmpPos2.Set px + MaxWidLenBack * TmpVec2.x, py + MaxWidLenBack * TmpVec2.y, pz + MaxWidLenBack * TmpVec2.z
            End If
            Set TmpVec2 = TmpPos1.Subtract(TmpPos2)
            TmpVec2.length = 1#
            If (Abs(TmpVec1.x - TmpVec2.x) <= 0.0001 And Abs(TmpVec1.y - TmpVec2.y) <= 0.0001 And Abs(TmpVec1.z - TmpVec2.z) <= 0.0001) Or _
                (Abs(TmpVec1.x + TmpVec2.x) <= 0.0001 And Abs(TmpVec1.y + TmpVec2.y) <= 0.0001 And Abs(TmpVec1.z + TmpVec2.z) <= 0.0001) Then
                CosAngle = 1#
            Else
                CosAngle = Abs(TmpVec1.x * TmpVec2.x + TmpVec1.y * TmpVec2.y + TmpVec1.z * TmpVec2.z)
            End If
            If Abs(CosAngle) <= 0.0001 Then
                WebAngle = 90#
            Else
                WebAngle = 180# * (Atn(Sqr(1 - CosAngle * CosAngle) / CosAngle)) / PI
            End If
        End If
    End If

    ' now compute the new cutting plane and normal in case of squared end
    If CutLength < ProjLength And CutDepth < ProjWidth Then
        CutDepth = ProjWidth
        ' TR#75014 ; do it based on which end is cut
        If WhichEnd = SPSMemberAxisStart Then
            ' use the max values either at top or bottom
            If BotMaxLenFlag = True Then
                TmpDist1 = MaxLenBot
            Else
                TmpDist1 = MaxLenTop
            End If
        Else
            ' use the min values either at top or bottom
            If BotMinLenFlag = True Then
                TmpDist1 = MinLenBot
            Else
                TmpDist1 = MinLenTop
            End If
        End If
        oPlaneRoot.Set oTopFacePos(1).x + TmpDist1 * MemVec.x, oTopFacePos(1).y + TmpDist1 * MemVec.y, oTopFacePos(1).z + TmpDist1 * MemVec.z
        oPlaneNormal.Set MemVec.x, MemVec.y, MemVec.z
    ElseIf Abs(CutDepth - ProjWidth) <= 0.0001 And Abs(CutLength - ProjLength) > 0.0001 Then
        ' TR#75014 ; do it based on which end is cut
        If WhichEnd = SPSMemberAxisStart Then
            ' use the max values either at top or bottom
            If BotMaxLenFlag = True Then
                TmpDist1 = MaxLenBot
            Else
                TmpDist1 = MaxLenTop
            End If
        Else
            ' use the min values either at top or bottom
            If BotMinLenFlag = True Then
                TmpDist1 = MinLenBot
            Else
                TmpDist1 = MinLenTop
            End If
        End If
        oPlaneRoot.Set oTopFacePos(1).x + TmpDist1 * MemVec.x, oTopFacePos(1).y + TmpDist1 * MemVec.y, oTopFacePos(1).z + TmpDist1 * MemVec.z
        oPlaneNormal.Set MemVec.x, MemVec.y, MemVec.z
    ElseIf Abs(CutLength - ProjLength) <= 0.0001 And Abs(CutDepth - ProjWidth) > 0.0001 Then
            If BotWidFlag = True Then
                TmpDist1 = MaxWidBot
            Else
                TmpDist1 = MaxWidTop
            End If
            TmpVec.Set oTopFacePos(2).x - oTopFacePos(1).x, oTopFacePos(2).y - oTopFacePos(1).y, oTopFacePos(2).z - oTopFacePos(1).z
            TmpVec.length = 1#
            oPlaneRoot.Set oTopFacePos(1).x + TmpDist1 * TmpVec.x, oTopFacePos(1).y + TmpDist1 * TmpVec.y, oTopFacePos(1).z + TmpDist1 * TmpVec.z
            oPlaneNormal.Set -TmpVec.x, -TmpVec.y, -TmpVec.z
    Else ' control should never come here ; both are cut to extreme
        GoTo ErrorHandler
    End If

    If SquaredEnd = True Then
        ' TR#77632 ; moved following lines from 'if' block
        Dim myplane As IJPlane
        Dim oIJGeometryMisc As IJGeometryMisc
        Set oIJGeometryMisc = New DGeomOpsMisc
        
        If oCutbackSurf Is Nothing Then
            Set myplane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, oPlaneRoot.x, oPlaneRoot.y, oPlaneRoot.z, oPlaneNormal.x, oPlaneNormal.y, oPlaneNormal.z) ' TR#77632 ; create a transient plane
            oIJGeometryMisc.CreateModelGeometryFromGType pResourceManager, myplane, Nothing, oCutbackSurf ' TR#77632
            Set myplane = Nothing
        ElseIf IsPlanar = True Then
            Set myplane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, oPlaneRoot.x, oPlaneRoot.y, oPlaneRoot.z, oPlaneNormal.x, oPlaneNormal.y, oPlaneNormal.z) ' TR#77632 ; create a transient plane' oCutbackSurf
            oIJGeometryMisc.ModifyModelGeometryFromGType myplane, oCutbackSurf ' TR#77632
            Set myplane = Nothing
        Else ' non-planar surface
            Dim tmpPlane As IJPlane

            Set tmpPlane = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, oPlaneRoot.x, oPlaneRoot.y, oPlaneRoot.z, oPlaneNormal.x, oPlaneNormal.y, oPlaneNormal.z)
            oIJGeometryMisc.ModifyModelGeometryFromGType tmpPlane, oCutbackSurf
            Set tmpPlane = Nothing
        End If
        Set oIJGeometryMisc = Nothing ' TR#77632
    End If

    For i = 1 To 4
        Set oTopFacePos(i) = Nothing
        Set oBotFacePos(i) = Nothing
    Next i
    Set pGeometryFactory = Nothing
    Set StartPos = Nothing
    Set EndPos = Nothing
    Set TmpPos1 = Nothing
    Set TmpPos2 = Nothing
    Set TmpPos3 = Nothing
    Set Line1Pos(0) = Nothing
    Set Line1Pos(1) = Nothing
    Set Line2Pos(0) = Nothing
    Set Line2Pos(1) = Nothing
    Set oPlaneRoot = Nothing
    Set oPlaneNormal = Nothing
    Set MemVec = Nothing
    Set TmpVec = Nothing
    Set TmpVec1 = Nothing
    Set TmpVec2 = Nothing
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub GetArcPointsAndCenter(ByVal RefPos As IJDPosition, ByVal Dir1 As IJDVector, ByVal Dir2 As IJDVector, ByVal Radius As Double, ByRef PtOnDir1 As IJDPosition, ByRef PtOnDir2 As IJDPosition, ByRef CentPt As IJDPosition)
    Const METHOD = "GetArcPointsAndCenter"
    On Error GoTo ErrorHandler
    Dim oCurve1 As IJCurve, oCurve2 As IJCurve
    Dim TmpVec3 As IJDVector, TmpVec4 As IJDVector, Angle As Double
    Dim NumInts As Long, IntPoints() As Double, overlap As Long
    Dim pGeometryFactory As New GeometryFactory
    Dim px As Double, py As Double, pz As Double
    Dim oLine As IJLine
    
    Set TmpVec3 = New DVector
    Set TmpVec4 = New DVector
    
    Set TmpVec3 = Dir1.Cross(Dir2)
    TmpVec3.length = 1#
    Angle = Dir1.Angle(Dir2, TmpVec3)
    If Angle > PI Then
        Angle = 2 * PI - Angle
    End If
    
    px = RefPos.x + Radius * Tan(Angle / 2#) * Dir1.x
    py = RefPos.y + Radius * Tan(Angle / 2#) * Dir1.y
    pz = RefPos.z + Radius * Tan(Angle / 2#) * Dir1.z
    PtOnDir1.Set px, py, pz

    px = RefPos.x + Radius * Tan(Angle / 2#) * Dir2.x
    py = RefPos.y + Radius * Tan(Angle / 2#) * Dir2.y
    pz = RefPos.z + Radius * Tan(Angle / 2#) * Dir2.z
    PtOnDir2.Set px, py, pz
    
    Set TmpVec4 = TmpVec3.Cross(Dir1) ' direction perpendicular to Dir1 and in the same plane
    TmpVec4.length = 1#
    Set oLine = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, PtOnDir1.x, PtOnDir1.y, PtOnDir1.z, PtOnDir1.x + LARGE_EDGE * TmpVec4.x, PtOnDir1.y + LARGE_EDGE * TmpVec4.y, PtOnDir1.z + LARGE_EDGE * TmpVec4.z)
    oLine.Infinite = True
    Set oCurve1 = oLine
    Set oLine = Nothing

    Set TmpVec4 = TmpVec3.Cross(Dir2) ' direction perpendicular to Dir2 and in the same plane
    TmpVec4.length = 1#
    Set oLine = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, PtOnDir2.x, PtOnDir2.y, PtOnDir2.z, PtOnDir2.x + LARGE_EDGE * TmpVec4.x, PtOnDir2.y + LARGE_EDGE * TmpVec4.y, PtOnDir2.z + LARGE_EDGE * TmpVec4.z)
    oLine.Infinite = True
    Set oCurve2 = oLine
    Set oLine = Nothing
    oCurve1.Intersect oCurve2, NumInts, IntPoints, overlap, ISECT_UNKNOWN
    If NumInts <= 0 Then
        GoTo ErrorHandler
    End If
    CentPt.Set IntPoints(0), IntPoints(1), IntPoints(2)
                
    Set oCurve1 = Nothing
    Set oCurve2 = Nothing
    Set TmpVec3 = Nothing
    Set TmpVec4 = Nothing
    Set pGeometryFactory = Nothing
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub GetFacePoints(oMemb As ISPSMemberPartPrismatic, TopBottom As Integer, Bounding As Boolean, oFacePos() As IJDPosition)
  Const MT = "GetFacePoints"
    On Error GoTo ErrorHandler
    
    Dim oPos As IJDPosition
    Dim oProfileBO As ISPSCrossSection
    Dim yOffset As Double, zOffset As Double, yOffsetCP As Double, zOffsetCP As Double
    Dim oMat As IJDT4x4
    Dim eX#, eY#, eZ#, sX#, sY#, sZ#, dblLength#
    Dim oVec As IJDVector
    Dim cp As Long, i As Integer, TopCPs(1 To 4) As Integer, BotCPs(1 To 4) As Integer
    Dim oSectionAttrbs As IJDAttributes
    Dim depth As Double, Width As Double, tWeb As Double, tFlange As Double

    Set oPos = New DPosition
    Set oVec = New DVector
    
    TopCPs(1) = 7
    TopCPs(2) = 9
    TopCPs(3) = 9
    TopCPs(4) = 7
    BotCPs(1) = 1
    BotCPs(2) = 3
    BotCPs(3) = 3
    BotCPs(4) = 1

    oMemb.Rotation.GetTransform oMat 'Get the member coordinate system
    If Not oMemb Is Nothing Then
        Set oProfileBO = oMemb.CrossSection
        If Not oProfileBO Is Nothing Then
            Dim oCrossection As IJCrossSection
            Set oCrossection = oProfileBO.definition
            If Not oCrossection Is Nothing Then
                Dim strSectionType As String
                strSectionType = oCrossection.Type
            End If
        End If
    End If
    
    oMemb.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    oVec.Set eX - sX, eY - sY, eZ - sZ
    dblLength = oVec.length
    cp = oProfileBO.CardinalPoint
    oProfileBO.GetCardinalPointOffset cp, yOffsetCP, zOffsetCP 'Returns Rad x and y of active CP which is member local -y and z.
    
    For i = 1 To 4
        If TopBottom = 1 Then ' top face required
            oProfileBO.GetCardinalPointOffset TopCPs(i), yOffset, zOffset 'Returns Rad x and y of CP=1 which is member local -y and z.
        Else
            oProfileBO.GetCardinalPointOffset BotCPs(i), yOffset, zOffset 'Returns Rad x and y of CP=1 which is member local -y and z.
        End If
        yOffset = yOffset - yOffsetCP ' get the y offset from current CP
        zOffset = zOffset - zOffsetCP ' get the z offset from current CP
        If i <= 2 Then ' get the points on the MemStart end
            oPos.Set 0, -1# * yOffset, zOffset ' point is now in the local coordinate system of the member. The yooset is multiplied by -1 because RAD + x is member local -Y.
        Else ' get the points on the MemEnd end
            oPos.Set dblLength, -1# * yOffset, zOffset ' point is now in the local coordinate system of the member. The yooset is multiplied by -1 because RAD + x is member local -Y.
        End If
        Set oFacePos(i) = oMat.TransformPosition(oPos)
    Next i

    If Bounding = False Then ' caller wants true shape instead of bounding
    Set oSectionAttrbs = oProfileBO.definition
    Select Case strSectionType
        Case "HSSC", "CS", "PIPE"
            ' just return with bounding top/bottom faces
        Case "W", "S", "HP", "M"
            ' just return with bounding top/bottom faces
        Case "L"
             'return the appropriate shape points
            tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
            If TopBottom = 1 Then
                oVec.Set oFacePos(2).x - oFacePos(1).x, oFacePos(2).y - oFacePos(1).y, oFacePos(2).z - oFacePos(1).z
                oVec.length = 1#
                oFacePos(2).Set oFacePos(1).x + tWeb * oVec.x, oFacePos(1).y + tWeb * oVec.y, oFacePos(1).z + tWeb * oVec.z
                oFacePos(3).Set oFacePos(4).x + tWeb * oVec.x, oFacePos(4).y + tWeb * oVec.y, oFacePos(4).z + tWeb * oVec.z
            Else
                ' just return with bounding points
            End If
        Case "C", "MC"
            ' just return with bounding top/bottom faces
        Case "WT", "MT", "ST"
             'return the appropriate shape points
            tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
            If TopBottom = 1 Then
                ' just return with bounding points
            Else
                oVec.Set oFacePos(2).x - oFacePos(1).x, oFacePos(2).y - oFacePos(1).y, oFacePos(2).z - oFacePos(1).z
                oVec.length = 1#
                Width = oFacePos(1).DistPt(oFacePos(2))
                oPos.Set oFacePos(1).x, oFacePos(1).y, oFacePos(1).z
                oFacePos(2).Set oPos.x + 0.5 * (Width + tWeb) * oVec.x, oPos.y + 0.5 * (Width + tWeb) * oVec.y, oPos.z + 0.5 * (Width + tWeb) * oVec.z
                oFacePos(1).Set oPos.x + 0.5 * (Width - tWeb) * oVec.x, oPos.y + 0.5 * (Width - tWeb) * oVec.y, oPos.z + 0.5 * (Width - tWeb) * oVec.z
                oPos.Set oFacePos(4).x, oFacePos(4).y, oFacePos(4).z
                oFacePos(3).Set oPos.x + 0.5 * (Width + tWeb) * oVec.x, oPos.y + 0.5 * (Width + tWeb) * oVec.y, oPos.z + 0.5 * (Width + tWeb) * oVec.z
                oFacePos(4).Set oPos.x + 0.5 * (Width - tWeb) * oVec.x, oPos.y + 0.5 * (Width - tWeb) * oVec.y, oPos.z + 0.5 * (Width - tWeb) * oVec.z
            End If
        Case "2L"
    
        Case "RS", "HSSR"
            ' just return with bounding top/bottom faces
        Case Else
            ' unknown type ; just return with bounding top/bottom faces
    End Select
    End If
    
    Set oSectionAttrbs = Nothing
    Set oProfileBO = Nothing
    Set oPos = Nothing
    Set oVec = Nothing
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
Private Function ComputeCentroidOfShape(ShpElms As IJElements, ByRef cgx As Double, ByRef cgy As Double, ByRef cgz As Double) As Integer
  Const MT = "ComputeCentroidOfShape"
    On Error GoTo ErrorHandler

    Dim elmCount As Integer, i As Integer, par As Double
    Dim myCurve As IJCurve, endCurve As IJCurve
    Dim sX1 As Double, sY1 As Double, sZ1 As Double, eX1 As Double, eY1 As Double, eZ1 As Double
    Dim refx As Double, refy As Double, refz As Double, midx As Double, midy As Double, midz As Double
    Dim density As Double, WtMomentX As Double, WtMomentY As Double, WtMomentZ As Double, wt As Double
    Dim px As Double, py As Double, pz As Double
    
    If ShpElms Is Nothing Or ((Not ShpElms Is Nothing) And ShpElms.count <= 0) Then
        ComputeCentroidOfShape = 1
        Exit Function
    End If
    density = 1#
    elmCount = ShpElms.count
    Set myCurve = ShpElms.Item(1)
    myCurve.EndPoints refx, refy, refz, px, py, pz
    Set myCurve = Nothing
    If elmCount > 1 Then
        Set myCurve = ShpElms.Item(2)
        myCurve.EndPoints sX1, sY1, sZ1, eX1, eY1, eZ1
        If (Abs(refx - sX1) <= 0.0001 And Abs(refy - sY1) <= 0.0001 And Abs(refz - sZ1) <= 0.0001) Or _
            (Abs(refx - eX1) <= 0.0001 And Abs(refy - eY1) <= 0.0001 And Abs(refz - eZ1) <= 0.0001) Then
            refx = px
            refy = py
            refz = pz
        Else ' don't do anything ; refx, refy, refz already store the starting point
        End If
        Set myCurve = Nothing
    End If
    For i = 1 To elmCount
        Set myCurve = ShpElms.Item(i)
        myCurve.EndPoints sX1, sY1, sZ1, eX1, eY1, eZ1
        If Abs(sX1 - eX1) <= 0.0001 And Abs(sY1 - eY1) <= 0.0001 And Abs(sZ1 - eZ1) <= 0.0001 Then
            ' closed shape ; get its centroid
            myCurve.Centroid midx, midy, midz
        Else
            myCurve.ParameterFRatio 0.5, par
            myCurve.Position par, midx, midy, midz
        End If
        wt = density * myCurve.length
        WtMomentX = wt * (midx - refx)
        WtMomentY = wt * (midy - refy)
        WtMomentZ = wt * (midz - refz)
        Set myCurve = Nothing
    Next i
    cgx = refx + WtMomentX / wt
    cgy = refy + WtMomentY / wt
    cgz = refz + WtMomentZ / wt
    ComputeCentroidOfShape = 0
    Exit Function

ErrorHandler:
    ComputeCentroidOfShape = 1
    HandleError MODULE, MT
End Function

Public Function IsFlangeWebCopeFormed(oSuppedPart As ISPSMemberPartPrismatic, oSuppingPlane1 As IJPlane, oSuppingPlane2 As IJPlane, FeaType As Integer) As Boolean
  Const MT = "IsFlangeWebCopeFormed"
    On Error GoTo ErrorHandler
    Dim ShpElms As IJElements, status As Integer, oLine As IJLine
    Dim px As Double, py As Double, pz As Double
    Dim oRootPos1 As IJDPosition, oRootPos2 As IJDPosition, TmpPos As IJDPosition, TmpPos2 As IJDPosition
    Dim oNormal1 As IJDVector, oNormal2 As IJDVector, ShpNorm As IJDVector, ShpNorm2 As IJDVector
    Dim oTopFacePos(1 To 4) As IJDPosition, oBotFacePos(1 To 4) As IJDPosition, MemStartPos As IJDPosition, MemEndPos As IJDPosition
    Dim BotTopDir As IJDVector, memDir As IJDVector, WebDir As IJDVector, DirCheckVec As IJDVector, NormVec As IJDVector
    Dim CosAngle As Double, Angle As Double, bIsFlangeWebCopeFormed As Boolean, Iter As Integer
    
    Set oRootPos1 = New DPosition
    Set oRootPos2 = New DPosition
    Set MemStartPos = New DPosition
    Set MemEndPos = New DPosition
    Set TmpPos = New DPosition
    Set TmpPos2 = New DPosition
    Set oNormal1 = New DVector
    Set oNormal2 = New DVector
    Set ShpNorm = New DVector
    Set ShpNorm2 = New DVector
    Set BotTopDir = New DVector
    Set memDir = New DVector
    Set WebDir = New DVector
    Set DirCheckVec = New DVector
    Set NormVec = New DVector
    Set oTopFacePos(1) = New DPosition
    Set oTopFacePos(2) = New DPosition
    Set oTopFacePos(3) = New DPosition
    Set oTopFacePos(4) = New DPosition
    Set oBotFacePos(1) = New DPosition
    Set oBotFacePos(2) = New DPosition
    Set oBotFacePos(3) = New DPosition
    Set oBotFacePos(4) = New DPosition

    oSuppingPlane1.GetBoundaries ShpElms
    status = ComputeCentroidOfShape(ShpElms, px, py, pz)
    If status = 0 Then ' success
    Else
        oSuppingPlane1.GetRootPoint px, py, pz
    End If
    Set ShpElms = Nothing
    oRootPos1.Set px, py, pz
    oSuppingPlane1.GetNormal px, py, pz
    oNormal1.Set px, py, pz
    oSuppingPlane2.GetBoundaries ShpElms
    status = ComputeCentroidOfShape(ShpElms, px, py, pz)
    If status = 0 Then ' success
    Else
        oSuppingPlane2.GetRootPoint px, py, pz
    End If
    Set ShpElms = Nothing
    oRootPos2.Set px, py, pz
    oSuppingPlane2.GetNormal px, py, pz
    oNormal2.Set px, py, pz
    oNormal1.length = 1#
    oNormal2.length = 1#
    
    Set ShpNorm = oNormal1.Cross(oNormal2)
    ShpNorm.length = 1#
    GetFacePoints oSuppedPart, 1, True, oTopFacePos
    GetFacePoints oSuppedPart, 2, True, oBotFacePos
    
    BotTopDir.Set oTopFacePos(1).x - oBotFacePos(1).x, oTopFacePos(1).y - oBotFacePos(1).y, oTopFacePos(1).z - oBotFacePos(1).z
    BotTopDir.length = 1#
    Set oLine = oSuppedPart
    oLine.GetStartPoint px, py, pz
    MemStartPos.Set px, py, pz
    oLine.GetEndPoint px, py, pz
    MemEndPos.Set px, py, pz
    Set oLine = Nothing
    memDir.Set MemStartPos.x - MemEndPos.x, MemStartPos.y - MemEndPos.y, MemStartPos.z - MemEndPos.z
    memDir.length = 1#
    Set WebDir = memDir.Cross(BotTopDir)
    WebDir.length = 1#
    
    bIsFlangeWebCopeFormed = True
    For Iter = 1 To 2
        If FeaType = 2 Then ' FlangeCope check
            If Iter = 1 Then
                NormVec.Set memDir.x, memDir.y, memDir.z
            Else
                NormVec.Set WebDir.x, WebDir.y, WebDir.z
            End If
            DirCheckVec.Set BotTopDir.x, BotTopDir.y, BotTopDir.z
        ElseIf FeaType = 1 Then ' WebCope check
            If Iter = 1 Then
                NormVec.Set BotTopDir.x, BotTopDir.y, BotTopDir.z
            Else
                NormVec.Set memDir.x, memDir.y, memDir.z
            End If
            DirCheckVec.Set WebDir.x, WebDir.y, WebDir.z
        Else ' unknown feature type to be checked
            GoTo ErrorHandler
        End If
        If bIsFlangeWebCopeFormed = True Then
            TmpPos.Set MemStartPos.x + ShpNorm.x, MemStartPos.y + ShpNorm.y, MemStartPos.z + ShpNorm.z
            Set TmpPos2 = ProjectPosToPlane(TmpPos, MemStartPos, NormVec)
            ShpNorm2.Set TmpPos2.x - MemStartPos.x, TmpPos2.y - MemStartPos.y, TmpPos2.z - MemStartPos.z
            ShpNorm2.length = 1#
            If (Abs(ShpNorm2.x - DirCheckVec.x) <= 0.0001 And Abs(ShpNorm2.y - DirCheckVec.y) <= 0.0001 And _
                Abs(ShpNorm2.z - DirCheckVec.z) <= 0.0001) Or (Abs(ShpNorm2.x + DirCheckVec.x) <= 0.0001 _
                And Abs(ShpNorm2.y + DirCheckVec.y) <= 0.0001 And Abs(ShpNorm2.z + DirCheckVec.z) <= 0.0001) Then
                bIsFlangeWebCopeFormed = True
            Else
                CosAngle = ShpNorm2.x * DirCheckVec.x + ShpNorm2.y * DirCheckVec.y + ShpNorm2.z * DirCheckVec.z
                If Abs(CosAngle) <= 0.0001 Then ' perpendicular
                    bIsFlangeWebCopeFormed = False
                Else
                    Angle = Atn(Sqr(1 - CosAngle * CosAngle) / Abs(CosAngle))
                    If Angle >= 0# And Angle <= 1# * PI / 180# Then
                        bIsFlangeWebCopeFormed = True
                    Else
                        bIsFlangeWebCopeFormed = False
                    End If
                End If
            End If
        End If
    Next Iter
    IsFlangeWebCopeFormed = bIsFlangeWebCopeFormed
    
    For Iter = 1 To 4
        Set oTopFacePos(Iter) = Nothing
        Set oBotFacePos(Iter) = Nothing
    Next Iter
    Exit Function
ErrorHandler:
    IsFlangeWebCopeFormed = False
    HandleError MODULE, MT
End Function

'Private Sub ComputeCopeSideEdgeForFlangeCope(oSuppedPart As ISPSMemberPartPrismatic, ByVal DrawMode As Integer, ByVal PlaneIntPt As IJDPosition, ByVal Shp1Vec As IJDVector, ByVal Shp2Vec As IJDVector, ByVal SideClr As Double, ByVal Clr1Vec As IJDVector, ByVal InsideClr As Double, ByVal Clr2Vec As IJDVector, ByVal PlanesConcave As Boolean, ByVal Radius As Double, ByVal RadiusType As Integer, ByRef CopeSide As Integer, ByRef CopeEdge As Integer)
'    Const METHOD = "ComputeCopeSideEdgeForFlangeCope"
'    On Error GoTo ErrorHandler
'    Dim px As Double, py As Double, pz As Double
'    Dim Circular As Boolean, SecTypeUnknown As Boolean, i As Integer, Iter As Integer
'    Dim oCurve1 As IJCurve, oCurve2 As IJCurve
'    Dim pGeometryFactory As New GeometryFactory
'    Dim idx As Integer
'    Dim TmpVec1 As IJDVector, TmpVec2 As IJDVector
'    Dim NumInts As Long, IntPoints() As Double, overlap As Long
'    Dim ProjStartPos As IJDPosition, ProjEndPos As IJDPosition, oTopFacePos(1 To 4) As IJDPosition, oBotFacePos(1 To 4) As IJDPosition
'    Dim minx As Double, miny As Double, minz As Double, maxx As Double, maxy As Double, maxz As Double
'    Dim oCmplx1 As ComplexString3d, oCmplx2 As ComplexString3d
'    Dim MemStartPos As IJDPosition, MemEndPos As IJDPosition, oLine As IJLine, TmpPos As IJDPosition, TmpPos2 As IJDPosition
'    Dim ProjLength As Double, ProjWidth As Double, TmpVec3 As IJDVector, TmpVec4 As IJDVector
'    Dim curveElms As IJElements, memElms As IJElements, LinePos(0 To 1) As IJDPosition
'    Dim MinLenDist As Double, MaxLenDist As Double, MinWidDist As Double, MaxWidDist As Double, TmpDist As Double, TmpDist2 As Double
'    Dim TmpVec5 As IJDVector, TmpVec6 As IJDVector, FirstIter As Boolean
'    Dim ProjIntSet(0 To 1) As Boolean, MinIntPt As IJDPosition, MaxIntPt As IJDPosition
'    Dim oProfile As IJCrossSection
'    Dim IsAtLeft As Boolean, TopFlange As Boolean, BotFlange As Boolean
'    Dim LeftMidRight(1 To 2) As Integer
'    Dim CopeLength As Double, CopeDepth As Double, MinWidOther As Double, MaxWidOther As Double
'
'    CopeSide = 0 ' initially set it to unknown cope type
'    CopeEdge = 0
'    Set oProfile = oSuppedPart.CrossSection.Definition
'    Circular = False
'    SecTypeUnknown = False
'    Select Case oProfile.Type
'        Case "HSSC", "CS", "PIPE"
'            Circular = True
'        Case "W", "S", "HP", "M"
'            TopFlange = True
'            BotFlange = True
'        Case "L"
'            TopFlange = False
'            BotFlange = True
'        Case "C", "MC"
'            TopFlange = True
'            BotFlange = True
'        Case "WT", "MT", "ST"
'            TopFlange = True
'            BotFlange = False
'        Case "2L"
'            'no top cope for 2L
'            TopFlange = False
'            BotFlange = True
'        Case "RS", "HSSR"
'            TopFlange = True
'            BotFlange = True
'        Case Else
'            SecTypeUnknown = True
'    End Select
'    Set oProfile = Nothing
'    If Circular = True Or SecTypeUnknown = True Then
'        Exit Sub
'    End If
'
'    Set oTopFacePos(1) = New DPosition
'    Set oTopFacePos(2) = New DPosition
'    Set oTopFacePos(3) = New DPosition
'    Set oTopFacePos(4) = New DPosition
'    Set oBotFacePos(1) = New DPosition
'    Set oBotFacePos(2) = New DPosition
'    Set oBotFacePos(3) = New DPosition
'    Set oBotFacePos(4) = New DPosition
'    Set MinIntPt = New DPosition
'    Set MaxIntPt = New DPosition
'    Set TmpVec1 = New DVector
'    Set TmpVec2 = New DVector
'    Set TmpVec3 = New DVector
'    Set TmpVec4 = New DVector
'    Set TmpVec5 = New DVector
'    Set TmpVec6 = New DVector
'    Set MemStartPos = New DPosition
'    Set MemEndPos = New DPosition
'    Set ProjStartPos = New DPosition
'    Set ProjEndPos = New DPosition
'    Set TmpPos = New DPosition
'    Set TmpPos2 = New DPosition
'    Set LinePos(0) = New DPosition
'    Set LinePos(1) = New DPosition
'
'    Set curveElms = New JObjectCollection
'    ' now draw the cope cutting shape taking into consideration clearances etc
'    DrawCornerCopeShape DrawMode, PlanesConcave, Radius, RadiusType, PlaneIntPt, Nothing, Shp1Vec, Shp2Vec, SideClr, Clr1Vec, InsideClr, Clr2Vec, curveElms
'
'    ' now get the rectangular outline of the member part and then get the intersection points
'    Set oLine = oSuppedPart
'    oLine.GetStartPoint px, py, pz
'    MemStartPos.Set px, py, pz
'    oLine.GetEndPoint px, py, pz
'    MemEndPos.Set px, py, pz
'    Set oLine = Nothing
'    Set TmpVec1 = Shp1Vec.Cross(Shp2Vec)
'    TmpVec1.length = 1#
'    Set ProjStartPos = ProjectPosToPlane(MemStartPos, PlaneIntPt, TmpVec1)
'    Set ProjEndPos = ProjectPosToPlane(MemEndPos, PlaneIntPt, TmpVec1)
'    TmpVec5.Set ProjEndPos.x - ProjStartPos.x, ProjEndPos.y - ProjStartPos.y, ProjEndPos.z - ProjStartPos.z
'    TmpVec5.length = 1#
'    Set TmpVec6 = TmpVec1.Cross(TmpVec5)
'    TmpVec6.length = 1#
'
'    GetFacePoints oSuppedPart, 1, False, oTopFacePos
'    GetFacePoints oSuppedPart, 2, False, oBotFacePos
'
'    ' initialize top, bot flags with proper values
'    LeftMidRight(1) = 0
'    LeftMidRight(2) = 0
'    For Iter = 1 To 2
'        If (Iter = 1 And TopFlange = False) Or (Iter = 2 And BotFlange = False) Then
'            GoTo CONTIN
'        End If
'
'      For i = 1 To 4
'        If Iter = 1 Then
'            Set TmpPos = ProjectPosToPlane(oTopFacePos(i), PlaneIntPt, TmpVec1)
'        Else
'            Set TmpPos = ProjectPosToPlane(oBotFacePos(i), PlaneIntPt, TmpVec1)
'        End If
'        TmpVec4.Set TmpPos.x - ProjStartPos.x, TmpPos.y - ProjStartPos.y, TmpPos.z - ProjStartPos.z
'        If TmpVec4.length > 0.0001 Then
'            TmpDist = TmpVec4.Dot(TmpVec5)
'        Else
'            TmpDist = 0#
'        End If
'        If MinLenDist > TmpDist Or i = 0 Then MinLenDist = TmpDist
'        If MaxLenDist < TmpDist Or i = 0 Then MaxLenDist = TmpDist
'        If TmpVec4.length > 0.0001 Then
'            TmpDist = TmpVec4.Dot(TmpVec6)
'        Else
'            TmpDist = 0#
'        End If
'        If MinWidDist > TmpDist Or i = 0 Then MinWidDist = TmpDist
'        If MaxWidDist < TmpDist Or i = 0 Then MaxWidDist = TmpDist
'      Next i
'
'    TmpPos.Set ProjStartPos.x + MinWidDist * TmpVec6.x, ProjStartPos.y + MinWidDist * TmpVec6.y, ProjStartPos.z + MinWidDist * TmpVec6.z
'    minx = TmpPos.x + MinLenDist * TmpVec5.x
'    miny = TmpPos.y + MinLenDist * TmpVec5.y
'    minz = TmpPos.z + MinLenDist * TmpVec5.z
'    TmpPos.Set ProjStartPos.x + MaxWidDist * TmpVec6.x, ProjStartPos.y + MaxWidDist * TmpVec6.y, ProjStartPos.z + MaxWidDist * TmpVec6.z
'    maxx = TmpPos.x + MaxLenDist * TmpVec5.x
'    maxy = TmpPos.y + MaxLenDist * TmpVec5.y
'    maxz = TmpPos.z + MaxLenDist * TmpVec5.z
'
'    TmpVec4.Set maxx - minx, maxy - miny, maxz - minz
'    ProjLength = Abs(TmpVec4.Dot(TmpVec5))
'    ProjWidth = Sqr(TmpVec4.length * TmpVec4.length - ProjLength * ProjLength)
'    TmpPos.Set maxx, maxy, maxz
'    LinePos(0).Set minx, miny, minz
'    LinePos(1).Set minx + ProjWidth * TmpVec6.x, miny + ProjWidth * TmpVec6.y, minz + ProjWidth * TmpVec6.z
'    ProjectPtOnLine TmpPos, LinePos, TmpPos2
'    TmpVec4.Set TmpPos2.x - minx, TmpPos2.y - miny, TmpPos2.z - minz
'    TmpVec4.length = 1#
'    LinePos(1).Set minx + ProjLength * TmpVec5.x, miny + ProjLength * TmpVec5.y, minz + ProjLength * TmpVec5.z
'    ProjectPtOnLine TmpPos, LinePos, TmpPos2
'    TmpVec3.Set TmpPos2.x - minx, TmpPos2.y - miny, TmpPos2.z - minz
'    TmpVec3.length = 1#
'
'    ' now determine whether minx, miny, minz is towards left or right
'    If Iter = 1 Then ' top face
'        LinePos(0).Set oTopFacePos(1).x, oTopFacePos(1).y, oTopFacePos(1).z
'        LinePos(1).Set oTopFacePos(2).x, oTopFacePos(2).y, oTopFacePos(2).z
'    Else
'        LinePos(0).Set oBotFacePos(1).x, oBotFacePos(1).y, oBotFacePos(1).z
'        LinePos(1).Set oBotFacePos(2).x, oBotFacePos(2).y, oBotFacePos(2).z
'    End If
'    TmpPos.Set minx, miny, minz
'    ProjectPtOnLine TmpPos, LinePos, TmpPos2
'    If Iter = 1 Then ' top face
'        TmpDist = TmpPos2.DistPt(oTopFacePos(1))
'        TmpDist2 = TmpPos2.DistPt(oTopFacePos(2))
'    Else
'        TmpDist = TmpPos2.DistPt(oBotFacePos(1))
'        TmpDist2 = TmpPos2.DistPt(oBotFacePos(2))
'    End If
'    If TmpDist <= TmpDist2 Then
'        ' pt minx, miny, minz is at left
'        IsAtLeft = True
'    Else
'        ' pt minx, miny, minz is at right
'        IsAtLeft = False
'    End If
'
'    ' now construct a member shape using min and max
'    ' tmpvec3 is along projected length, tmpvec4 is along projected width
'    Set memElms = New JObjectCollection
'    Set oLine = Nothing
'    Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, minx, miny, minz, TmpVec3.x, TmpVec3.y, TmpVec3.z, ProjLength)
'    memElms.Add oLine
'    Set oLine = Nothing
'    Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, minx + ProjLength * TmpVec3.x, miny + ProjLength * TmpVec3.y, minz + ProjLength * TmpVec3.z, TmpVec4.x, TmpVec4.y, TmpVec4.z, ProjWidth)
'    memElms.Add oLine
'    Set oLine = Nothing
'    Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, maxx, maxy, maxz, -1# * TmpVec3.x, -1# * TmpVec3.y, -1# * TmpVec3.z, ProjLength)
'    memElms.Add oLine
'    Set oLine = Nothing
'    Set oLine = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, minx + ProjWidth * TmpVec4.x, miny + ProjWidth * TmpVec4.y, minz + ProjWidth * TmpVec4.z, -1# * TmpVec4.x, -1# * TmpVec4.y, -1# * TmpVec4.z, ProjWidth)
'    memElms.Add oLine
'    Set oLine = Nothing
'    Set oCmplx2 = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, memElms)
'    Set oCmplx1 = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
'    GetIntersectionPointsOfCurves oCmplx1, oCmplx2, NumInts, IntPoints
'    Set oCmplx2 = Nothing
'    memElms.Clear
'    Set memElms = Nothing
'
'    MinLenDist = ProjLength
'    MaxLenDist = 0
'    MinWidDist = ProjWidth
'    MaxWidDist = 0
'    ProjIntSet(0) = False
'    ProjIntSet(1) = False
'    For i = 1 To NumInts
'        TmpPos.Set IntPoints((i - 1) * 3), IntPoints((i - 1) * 3 + 1), IntPoints((i - 1) * 3 + 2)
'
'        ' first in projlength direction
'        LinePos(0).Set minx, miny, minz
'        LinePos(1).Set minx + ProjLength * TmpVec3.x, miny + ProjLength * TmpVec3.y, minz + ProjLength * TmpVec3.z
'        ProjectPtOnLine TmpPos, LinePos, TmpPos2
'        TmpDist = Sqr((TmpPos2.x - minx) * (TmpPos2.x - minx) + (TmpPos2.y - miny) * (TmpPos2.y - miny) + (TmpPos2.z - minz) * (TmpPos2.z - minz))
'        If TmpDist <= 0.0001 Then
'            ProjIntSet(0) = True
'            MinIntPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'        ElseIf Abs(TmpDist - ProjLength) <= 0.0001 Then
'            ProjIntSet(1) = True
'            MaxIntPt.Set TmpPos.x, TmpPos.y, TmpPos.z
'        End If
'        If MinLenDist > TmpDist Then MinLenDist = TmpDist
'        If MaxLenDist < TmpDist Then MaxLenDist = TmpDist
'
'        ' now in projwidth direction
'        LinePos(1).Set minx + ProjWidth * TmpVec4.x, miny + ProjWidth * TmpVec4.y, minz + ProjWidth * TmpVec4.z
'        ProjectPtOnLine TmpPos, LinePos, TmpPos2
'        TmpDist = Sqr((TmpPos2.x - minx) * (TmpPos2.x - minx) + (TmpPos2.y - miny) * (TmpPos2.y - miny) + (TmpPos2.z - minz) * (TmpPos2.z - minz))
'        If MinWidDist > TmpDist Then MinWidDist = TmpDist
'        If MaxWidDist < TmpDist Then MaxWidDist = TmpDist
'    Next i
'    ' now do the same for PlaneIntPoint in ProjLength dir
'    LinePos(0).Set minx, miny, minz
'    LinePos(1).Set minx + ProjLength * TmpVec3.x, miny + ProjLength * TmpVec3.y, minz + ProjLength * TmpVec3.z
'    ProjectPtOnLine PlaneIntPt, LinePos, TmpPos2
'    TmpDist = Sqr((TmpPos2.x - minx) * (TmpPos2.x - minx) + (TmpPos2.y - miny) * (TmpPos2.y - miny) + (TmpPos2.z - minz) * (TmpPos2.z - minz))
'    If MinLenDist > TmpDist Then MinLenDist = TmpDist
'    If MaxLenDist < TmpDist Then MaxLenDist = TmpDist
'    LinePos(1).Set minx + ProjWidth * TmpVec4.x, miny + ProjWidth * TmpVec4.y, minz + ProjWidth * TmpVec4.z
'    ProjectPtOnLine PlaneIntPt, LinePos, TmpPos2
'    TmpDist = Sqr((TmpPos2.x - minx) * (TmpPos2.x - minx) + (TmpPos2.y - miny) * (TmpPos2.y - miny) + (TmpPos2.z - minz) * (TmpPos2.z - minz))
'    If MinWidDist > TmpDist Then MinWidDist = TmpDist
'    If MaxWidDist < TmpDist Then MaxWidDist = TmpDist
'
'    If Abs(MaxLenDist - MinLenDist) <= 0.0001 Then
'        CopeLength = MaxLenDist
'    Else
'        CopeLength = MaxLenDist - MinLenDist
'    End If
'    If Abs(MaxWidDist - MinWidDist) <= 0.0001 Then
'        CopeDepth = MaxWidDist
'    Else
'        CopeDepth = MaxWidDist - MinWidDist
'    End If
'
'    If CopeLength < ProjLength And CopeDepth < ProjWidth Then
'        If Abs(MinWidDist) <= 0.0001 Then
'            If IsAtLeft = True Then
'                    LeftMidRight(Iter) = 1 ' top/bot-left
'            Else
'                    LeftMidRight(Iter) = 2 ' top/bot-right
'            End If
'        ElseIf Abs(MaxWidDist - ProjWidth) <= 0.0001 Then
'            If IsAtLeft = True Then
'                LeftMidRight(Iter) = 2 ' top/bot-right
'            Else
'                LeftMidRight(Iter) = 1 ' top/bot-left
'            End If
'        Else ' min and max width lines are between the extremes
'            MinWidOther = ProjWidth - MaxWidDist
'            MaxWidOther = ProjWidth - MinWidDist
'            If MinWidDist <= MinWidOther Then
'                ' cope towards minx, miny, minz
'                If IsAtLeft = True Then
'                    LeftMidRight(Iter) = 1 ' top/bot-left
'                Else
'                    LeftMidRight(Iter) = 2 ' top/bot-right
'                End If
'            Else
'                ' cope towards other end of the edge
'                If IsAtLeft = True Then
'                    LeftMidRight(Iter) = 2 ' top/bot-right
'                Else
'                    LeftMidRight(Iter) = 1 ' top/bot-left
'                End If
'            End If
'        End If
'    ElseIf Abs(CopeDepth - ProjWidth) <= 0.0001 And Abs(CopeLength - ProjLength) > 0.0001 Then
'            LeftMidRight(Iter) = 3 ' top/bot-both left & right
'    ElseIf Abs(CopeLength - ProjLength) <= 0.0001 And Abs(CopeDepth - ProjWidth) > 0.0001 Then
'        Dim LowerEndCut As Boolean
'        ' try to figure out which part has to be removed
'        If ProjIntSet(0) = True Then
'            TmpDist = Sqr((MinIntPt.x - minx) * (MinIntPt.x - minx) + (MinIntPt.y - miny) * (MinIntPt.y - miny) + (MinIntPt.z - minz) * (MinIntPt.z - minz))
'            If TmpDist <= 0.5 * ProjWidth Then
'                LowerEndCut = True
'            Else
'                LowerEndCut = False
'            End If
'        ElseIf ProjIntSet(1) = True Then
'            TmpDist = Sqr((MaxIntPt.x - maxx) * (MaxIntPt.x - maxx) + (MaxIntPt.y - maxy) * (MaxIntPt.y - maxy) + (MaxIntPt.z - maxz) * (MaxIntPt.z - maxz))
'            If TmpDist > 0.5 * ProjWidth Then
'                LowerEndCut = True
'            Else
'                LowerEndCut = False
'            End If
'        Else
'            ' determine based on planeintpt
'            LinePos(0).Set minx, miny, minz
'            LinePos(0).Set minx + ProjLength * TmpVec3.x, miny + ProjLength * TmpVec3.y, minz + ProjLength * TmpVec3.z
'            ProjectPtOnLine PlaneIntPt, LinePos, TmpPos
'            TmpDist = Sqr((TmpPos.x - PlaneIntPt.x) * (TmpPos.x - PlaneIntPt.x) + (TmpPos.y - PlaneIntPt.y) * (TmpPos.y - PlaneIntPt.y) + (TmpPos.z - PlaneIntPt.z) * (TmpPos.z - PlaneIntPt.z))
'            If TmpDist <= 0.5 * ProjWidth Then
'                LowerEndCut = True
'            Else
'                LowerEndCut = False
'            End If
'        End If
'        If LowerEndCut = True Then
'            If IsAtLeft = True Then
'                LeftMidRight(Iter) = 1 ' top/bot-left
'            Else
'                LeftMidRight(Iter) = 2 ' top/bot-right
'            End If
'        Else
'            If IsAtLeft = True Then
'                LeftMidRight(Iter) = 2 ' top/bot-right
'            Else
'                LeftMidRight(Iter) = 1 ' top/bot-left
'            End If
'        End If
'    Else ' control should never come here
'        GoTo ErrorHandler
'    End If
'CONTIN:
'    Next Iter
'
'    ' now based on top, bottom flags determine the copeside, copeedge
'    ' 1 - left ; 2 - right ; 3 - full
'    If (LeftMidRight(1) = 0 And LeftMidRight(2) <> 0) Or (LeftMidRight(1) <> 0 And LeftMidRight(2) = 0) Then
'        If LeftMidRight(1) <> 0 Then
'            CopeSide = LeftMidRight(1)
'            CopeEdge = 1
'        ElseIf LeftMidRight(2) <> 0 Then
'            CopeSide = LeftMidRight(2)
'            CopeEdge = 2
'        End If
'    ElseIf LeftMidRight(1) <> 0 And LeftMidRight(2) <> 0 Then
'            Select Case LeftMidRight(1)
'                    Case 1
'                        Select Case LeftMidRight(2)
'                            Case 1 ' TL & BL
'                                CopeSide = 9
'                            Case 2 ' TL and BR
'                                CopeSide = 11
'                            Case 3 ' TL and BF
'                                CopeSide = 12
'                        End Select
'                    Case 2
'                        Select Case LeftMidRight(2)
'                            Case 1 ' TR & BL
'                                CopeSide = 17
'                            Case 2 ' TR and BR
'                                CopeSide = 19
'                            Case 3 ' TR and BF
'                                CopeSide = 20
'                        End Select
'                    Case 3
'                        Select Case LeftMidRight(2)
'                            Case 1 ' TF & BL
'                                CopeSide = 21
'                            Case 2 ' TF and BR
'                                CopeSide = 23
'                            Case 3 ' TF and BF
'                                CopeSide = 24
'                        End Select
'            End Select
'    End If
'MsgBox "CopeSide = " & CopeSide
'
'    For i = 1 To 4
'        Set oTopFacePos(i) = Nothing
'        Set oBotFacePos(i) = Nothing
'    Next i
'    Set oCmplx1 = Nothing
'    curveElms.Clear
'    Set curveElms = Nothing
'    Set MemStartPos = Nothing
'    Set MemEndPos = Nothing
'    Set ProjStartPos = Nothing
'    Set ProjEndPos = Nothing
'    Set TmpVec1 = Nothing
'    Set TmpVec2 = Nothing
'    Set TmpVec3 = Nothing
'    Set TmpVec4 = Nothing
'    Set TmpVec5 = Nothing
'    Set TmpVec6 = Nothing
'    Set pGeometryFactory = Nothing
'    Set TmpPos = Nothing
'    Set TmpPos2 = Nothing
'    Set LinePos(0) = Nothing
'    Set LinePos(1) = Nothing
'    Set MinIntPt = Nothing
'    Set MaxIntPt = Nothing
'    Exit Sub
'ErrorHandler:
'    HandleError MODULE, METHOD
'End Sub

Public Sub GetCrossSectionPoints(oMemb As ISPSMemberPartPrismatic, iPort As SPSMemberAxisPortIndex, Bounding As Boolean, PosCount As Integer, oCSPos() As IJDPosition) ' TR#75015 ; added argument to get bounding points
  Const MT = "GetCrossSectionPoints"
    On Error GoTo ErrorHandler
    
    Dim oSectionAttrbs As IJDAttributes
    Dim depth As Double, tFlange As Double, tWeb As Double, bFlange As Double
    Dim oPos() As IJDPosition
    Dim i As Integer
    Dim oMat As IJDT4x4
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#, nx#, ny#, nz#
    Dim oProfile As IJCrossSection
    Dim bTob#
    Dim xVec As New DVector
    Dim membLength As Double
   
    If Not oMemb Is Nothing Then
        Dim OSPSCrossSection As ISPSCrossSection
        Set OSPSCrossSection = oMemb.CrossSection
        If Not OSPSCrossSection Is Nothing Then
            Set oProfile = OSPSCrossSection.definition
            If Not oProfile Is Nothing Then
                Dim strSectionType As String
                strSectionType = oProfile.Type
            End If
        End If
    End If
    Set oSectionAttrbs = oMemb.CrossSection.definition
    
    'for now assume that the section type is W
    'get section attributes
    depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    On Error Resume Next ' some sections don't support the interfaces below
    tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
    tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
    
    On Error GoTo ErrorHandler
    membLength = oMemb.Axis.length
    oMemb.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    
    'represents x axis of member
    xVec.Set eX - sX, eY - sY, eZ - sZ
    xVec.length = 1 'normalize the vector
    
    ' TR#75015 ; get the bounding rectangle / rnage box, if asked for
    If Bounding = True Then
        PosCount = 4
        ReDim oPos(1 To PosCount)
        For i = 1 To PosCount
            Set oPos(i) = New DPosition
        Next i
        
        'oPos(1) is CP7 and going clockwise
        oPos(1).Set -bFlange / 2, depth / 2, 0
        oPos(2).Set -oPos(1).x, oPos(1).y, 0
        oPos(3).Set oPos(2).x, -oPos(2).y, 0
        oPos(4).Set oPos(1).x, -oPos(1).y, 0
        GoTo GLOBALPOINTS
    End If
    ' TR#75015 ends here
    
    Select Case strSectionType
        Case "HSSC", "CS", "PIPE", "RS", "HSSR"
            PosCount = 4
            ReDim oPos(1 To PosCount)
            For i = 1 To PosCount
                Set oPos(i) = New DPosition
            Next i

            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, -oPos(2).y, 0
            oPos(4).Set oPos(1).x, -oPos(1).y, 0
        
        Case "W", "S", "HP", "M"
            PosCount = 12

            ReDim oPos(1 To PosCount)
            For i = 1 To PosCount
                Set oPos(i) = New DPosition
            Next i

            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, oPos(2).y - tFlange, 0
            oPos(4).Set tWeb / 2, oPos(3).y, 0
            oPos(5).Set oPos(4).x, -oPos(4).y, 0
            oPos(6).Set oPos(3).x, -oPos(3).y, 0
            oPos(7).Set oPos(2).x, -oPos(2).y, 0
            oPos(8).Set oPos(1).x, -oPos(1).y, 0
            oPos(9).Set -oPos(6).x, oPos(6).y, 0
            oPos(10).Set -oPos(5).x, oPos(5).y, 0
            oPos(11).Set oPos(10).x, -oPos(10).y, 0
            oPos(12).Set oPos(9).x, -oPos(9).y, 0
        
        Case "L"
            PosCount = 6
            ReDim oPos(1 To PosCount)
            For i = 1 To PosCount
                Set oPos(i) = New DPosition
            Next i

            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set oPos(1).x + tWeb, oPos(1).y, 0 ' TR#75015 ; small bug
            oPos(3).Set oPos(2).x, -depth / 2 + tFlange, 0
            oPos(4).Set -oPos(1).x, oPos(3).y, 0
            oPos(5).Set oPos(4).x, -oPos(1).y, 0
            oPos(6).Set oPos(1).x, -oPos(1).y, 0
        
        Case "C", "MC"
            PosCount = 8
            ReDim oPos(1 To PosCount)
            For i = 1 To PosCount
                Set oPos(i) = New DPosition
            Next i

            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, oPos(2).y - tFlange, 0
            oPos(4).Set oPos(1).x + tWeb, oPos(3).y, 0 ' TR#75015 ; small bug
            oPos(5).Set oPos(4).x, -oPos(4).y, 0
            oPos(6).Set oPos(3).x, -oPos(3).y, 0
            oPos(7).Set oPos(2).x, -oPos(2).y, 0
            oPos(8).Set oPos(1).x, -oPos(1).y, 0

        Case "WT", "MT", "ST"
            PosCount = 8
            ReDim oPos(1 To PosCount)
            For i = 1 To PosCount
                Set oPos(i) = New DPosition
            Next i
            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, oPos(2).y - tFlange, 0
            oPos(4).Set tWeb / 2, oPos(3).y, 0
            oPos(5).Set oPos(4).x, -oPos(2).y, 0 ' TR#75015 ; small bug
            
            oPos(6).Set -oPos(5).x, oPos(5).y, 0
            oPos(7).Set -oPos(4).x, oPos(4).y, 0
            oPos(8).Set -oPos(3).x, oPos(3).y, 0
        Case "2L"
            bTob = oSectionAttrbs.CollectionOfAttributes("IJUA2L").Item("bb").Value
            PosCount = 8
            ReDim oPos(1 To PosCount)
            For i = 1 To PosCount
                Set oPos(i) = New DPosition
            Next i
            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -(tWeb / 2 + bTob / 2), depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, -(depth / 2 - tFlange), 0
            oPos(4).Set bFlange / 2, oPos(3).y, 0
            oPos(5).Set oPos(4).x, -depth / 2, 0
            oPos(6).Set -oPos(5).x, oPos(5).y, 0
            oPos(7).Set -oPos(4).x, oPos(4).y, 0
            oPos(8).Set -oPos(3).x, oPos(3).y, 0
        Case Else
            'unknown cross section. Create based on range box
            PosCount = 4
            ReDim oPos(1 To PosCount)
            For i = 1 To PosCount
                Set oPos(i) = New DPosition
            Next i

            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, -oPos(2).y, 0
            oPos(4).Set oPos(1).x, -oPos(1).y, 0
    End Select

GLOBALPOINTS: ' TR#75015
    ReDim oCSPos(0 To PosCount - 1) As IJDPosition
    For i = 0 To PosCount - 1
        'transform Position to global
        Set oCSPos(i) = TransformPosToGlobal(oMemb, oPos(i + 1))
        If iPort = SPSMemberAxisEnd Then
            ' now translate to corresponding point at other end
            oCSPos(i).Get sX, sY, sZ
            eX = sX + membLength * xVec.x
            eY = sY + membLength * xVec.y
            eZ = sZ + membLength * xVec.z
            oCSPos(i).Set eX, eY, eZ
        End If
        Set oPos(i + 1) = Nothing
    Next i
    
    Set xVec = Nothing
    Set oSectionAttrbs = Nothing
    Set oProfile = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub



'Private Sub MakeLineFinite(ByRef InfLine() As IJDPosition)
'Dim dirx As Double, diry As Double, dirz As Double, length As Double, mx As Double, my As Double, mz As Double, px As Double, py As Double, pz As Double
'
'dirx = InfLine(1).x - InfLine(0).x
'diry = InfLine(1).y - InfLine(0).y
'dirz = InfLine(1).z - InfLine(0).z
'length = Sqr(dirx * dirx + diry * diry + dirz * dirz)
'dirx = dirx / length
'diry = diry / length
'dirz = dirz / length
'
'' now get the mid point of the line
'mx = 0.5 * (InfLine(0).x + InfLine(1).x)
'my = 0.5 * (InfLine(0).y + InfLine(1).y)
'mz = 0.5 * (InfLine(0).z + InfLine(1).z)
'
'' now project the mid point on either side to finite length
'px = mx - LARGE_EDGE * dirx
'py = my - LARGE_EDGE * diry
'pz = mz - LARGE_EDGE * dirz
'InfLine(0).Set px, py, pz
'
'px = mx + LARGE_EDGE * dirx
'py = my + LARGE_EDGE * diry
'pz = mz + LARGE_EDGE * dirz
'InfLine(1).Set px, py, pz
'
'End Sub

' TR#76302 ; added a new function to determine whether a surface cuts both ends of a member
Public Function DoesNotCutBothEnds(oSuppedPart As ISPSMemberPartPrismatic, intElms As IJElements) As Boolean
  Const MT = "DoesNotCutBothEnds"
    On Error GoTo ErrorHandler
    Dim bNotBothEnds As Boolean
    Dim IntLine As Line3d, InterSecPoint As IJPoint
    Dim i As Integer, MemVec As IJDVector, DirVec As IJDVector, LinePos(1) As IJDPosition
    Dim px1 As Double, py1 As Double, pz1 As Double
    Dim px2 As Double, py2 As Double, pz2 As Double
    Dim midx As Double, midy As Double, midz As Double
    Dim oLine As IJLine, FirstEnd As SPSMemberAxisPortIndex, SomeEnd As SPSMemberAxisPortIndex, EleType As Integer
    Dim OrgPos As IJDPosition, ProjPos As IJDPosition
    Dim oCurve As IJCurve, par As Double
    
    bNotBothEnds = True
    If Not intElms Is Nothing Then
        If intElms.count > 0 Then
            Set MemVec = New DVector
            Set DirVec = New DVector
            Set LinePos(0) = New DPosition
            Set LinePos(1) = New DPosition
            Set OrgPos = New DPosition
            Set ProjPos = New DPosition
            Set oLine = oSuppedPart
            oLine.GetStartPoint px1, py1, pz1
            oLine.GetEndPoint px2, py2, pz2
            midx = (px1 + px2) * 0.5
            midy = (py1 + py2) * 0.5
            midz = (pz1 + pz2) * 0.5
            oLine.GetDirection px1, py1, pz1
            MemVec.Set px1, py1, pz1
            MemVec.length = 1#
            Set oLine = Nothing
            If TypeOf intElms.Item(1) Is IJPoint Then
                EleType = 1 ' IJPoint
            Else
                EleType = 2 ' IJCurve
            End If
            For i = 1 To intElms.count
                If EleType = 1 Then
                    Set InterSecPoint = intElms.Item(i)
                    InterSecPoint.GetPoint px1, py1, pz1
                    ProjPos.Set px1, py1, pz1
                Else
                    Set oCurve = intElms.Item(i)
                    oCurve.ParameterFRatio 0.5, par
                    oCurve.Position par, px1, py1, pz1
                    ProjPos.Set px1, py1, pz1
                    Set oCurve = Nothing
                End If
                DirVec.Set ProjPos.x - midx, ProjPos.y - midy, ProjPos.z - midz
                DirVec.length = 1#
                If MemVec.Dot(DirVec) > 0 Then
                    SomeEnd = SPSMemberAxisEnd
                ElseIf MemVec.Dot(DirVec) < 0 Then
                        SomeEnd = SPSMemberAxisStart
                Else ' exactly perpendicular to member axis
                        SomeEnd = SPSMemberAxisAlong
                End If
                If i = 1 Or FirstEnd = SPSMemberAxisAlong Then
                    FirstEnd = SomeEnd
                ElseIf FirstEnd <> SPSMemberAxisAlong And SomeEnd <> SPSMemberAxisAlong And FirstEnd <> SomeEnd Then
                    bNotBothEnds = False
                    Exit For
                End If
            Next i
            Set MemVec = Nothing
            Set DirVec = Nothing
            Set LinePos(0) = Nothing
            Set LinePos(1) = Nothing
            Set ProjPos = Nothing
            Set OrgPos = Nothing
        End If
    End If
    DoesNotCutBothEnds = bNotBothEnds
    Exit Function
    
ErrorHandler:
    DoesNotCutBothEnds = False
    HandleError MODULE, MT
End Function
