VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuiltUpPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements IJDMfgSystemMarkingRule
Implements IJDMfgSystemMarkingRule2

Private Const MODULE = "MfgMemberMarking.BuiltUpPC"
Private Const IID_IJMfgSystemMarkInfo As Variant = "{CE13B4E9-C6B0-45D6-99F4-0981D870274E}"

' A couple of constants that decide how big and where fitting marks are created
Private Const PI As Double = 3.14159265358979

Private Const dFittingMarkLength As Double = 0.02
Private Const dFittingMarkDistanceFromStartOfWeb As Double = 0.5
    
Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    Helpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    Helpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function
'  Create three marking lines for the projection of the web left and the web right onto the flange face
'  The first marking line will result from the projection of the webleft or web right edge onto the flange surface
'  The other two marking lines will be created to align the flange and web.

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "BuiltUpPC: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateBuiltUpPCMark Part, JXSEC_WEB_LEFT, JXSEC_TOP, oGeomCol3d
    CreateBuiltUpPCMark Part, JXSEC_WEB_RIGHT, JXSEC_TOP, oGeomCol3d

    CreateBuiltUpPCMark Part, JXSEC_WEB_LEFT, JXSEC_BOTTOM, oGeomCol3d
    CreateBuiltUpPCMark Part, JXSEC_WEB_RIGHT, JXSEC_BOTTOM, oGeomCol3d
    
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2002, , "RULES")
End Function

Private Function IJDMfgSystemMarkingRule2_CreateAtTheEnd(ByVal oProfilePart As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

    Dim i As Integer, j As Integer
    Dim lGeomFaceId As Long, lGeomCount As Long
    Dim dLowestX As Double, dHighestX As Double
    Dim dWebLowestY As Double, dWebHighestY As Double, dFLowestY As Double, dFHighestY As Double
    Dim dFlangeWidth As Double, dFlangeThickness As Double
    Dim oGeom2d As IJMfgGeom2d, oFirstGeom2d As IJMfgGeom2d
    Dim oGeomCol2d As IJMfgGeomCol2d
    Dim oMfgPart As IJMfgProfilePart
    Dim eGeomType As StrMfgGeometryType
    Dim oHelperSupport As IJMfgRuleHelpersSupport
    Dim lVerticalMarks As Long
    Dim dVerticalMarksX() As Double, dVerticalMarksY() As Double

    Set oHelperSupport = New MfgRuleHelpersSupport
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfilePart
    
    If TypeOf oProfilePart Is IJProfileER Then
        Exit Function
    End If
    
    Dim oMfgProfileWrapper As MfgRuleHelpers.MfgProfilePartHlpr
    If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
        Set oMfgProfileWrapper = New MfgRuleHelpers.MfgProfilePartHlpr
        Set oMfgProfileWrapper.object = oMfgPart
    Else
        Exit Function
    End If
            
    Set oGeomCol2d = oMfgPart.FinalGeometriesAfterProcess2D
    
    CreateFlangeElevation oProfilePart, UpSide, oGeomCol2d, JXSEC_TOP
    CreateFlangeElevation oProfilePart, UpSide, oGeomCol2d, JXSEC_BOTTOM
    
    UpdateBuiltUpPCMarks oProfilePart, UpSide, oGeomCol2d, JXSEC_TOP
    UpdateBuiltUpPCMarks oProfilePart, UpSide, oGeomCol2d, JXSEC_BOTTOM
    
End Function
Private Function CreateFlangeElevation(ByVal oProfilePart As Object, ByVal UpSide As Long, ByVal oGeomCol2d As IJMfgGeomCol2d, _
                        ByVal lFlangeFaceID As Long)

    Dim i As Long, j As Long, lGeomCount As Long, lVerticalMarks As Long, lGeomFaceId As Long
    Dim dFlangeWidth As Double, dFlangeThickness As Double
    Dim dLowestX As Double, dHighestX As Double
    Dim dWebLowestY As Double, dWebHighestY As Double, dFLowestY As Double, dFHighestY As Double
    Dim dLowerLeftX As Double, dLowerLeftY As Double
    Dim dLowerRightX As Double, dLowerRightY As Double
    Dim dUpperLeftX As Double, dUpperLeftY As Double
    Dim dUpperRightX As Double, dUpperRightY As Double
    Dim oGeom2d As IJMfgGeom2d, oFirstGeom2d As IJMfgGeom2d
    Dim oProfilePartSupport As IJProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Dim eGeomType As StrMfgGeometryType
    Dim oHelperSupport As IJMfgRuleHelpersSupport
    Dim dVerticalMarksX() As Double, dVerticalMarksY() As Double
    Dim lGeomTypes() As Double

    Set oHelperSupport = New MfgRuleHelpersSupport
    
    Set oProfilePartSupport = New ProfilePartSupport
    Set oPartSupport = oProfilePartSupport
    Set oPartSupport.Part = oProfilePart
    
    oProfilePartSupport.GetFlangeWidth dFlangeWidth
    
    If lFlangeFaceID = JXSEC_TOP Then
        oProfilePartSupport.GetThickness TopFlange, dFlangeThickness
    Else
        oProfilePartSupport.GetThickness BottomFlange, dFlangeThickness
    End If
    
    lGeomCount = oGeomCol2d.Getcount
    
    lVerticalMarks = 0
    
    dLowestX = 1E+30
    dHighestX = -1E+31
    dWebLowestY = 1E+30
    dWebHighestY = -1E+31
    dFLowestY = 1E+30
    dFHighestY = -1E+31
    
    If ((dFlangeWidth > 0.001) And (dFlangeThickness > 0.001)) Then
        For i = 1 To lGeomCount
            Set oGeom2d = oGeomCol2d.GetGeometry(i)
            eGeomType = oGeom2d.GetGeometryType
            lGeomFaceId = oGeom2d.FaceId
            
            If ((eGeomType = STRMFG_OUTER_CONTOUR) Or (eGeomType = STRMFG_INNER_CONTOUR)) Then
                If (oGeom2d.GetProfileCell = Web) Then
                    If oFirstGeom2d Is Nothing Then
                        Set oFirstGeom2d = oGeom2d
                    End If
                End If
                                
                Dim dMinX As Double, dMinY As Double, dMinZ As Double
                Dim dMaxX As Double, dMaxY As Double, dMaxZ As Double
                
                'Get the Part Range
                oHelperSupport.GetRange oGeom2d, dMinX, dMinY, dMinZ, dMaxX, dMaxY, dMaxZ
                
                If (lGeomFaceId = lFlangeFaceID) Then
                    ReDim Preserve dVerticalMarksX(lVerticalMarks + 2)
                    ReDim Preserve dVerticalMarksY(lVerticalMarks + 2)
                    dVerticalMarksX(lVerticalMarks) = dMinX
                    dVerticalMarksY(lVerticalMarks) = (dMinY + dMaxY) / 2
                    dVerticalMarksX(lVerticalMarks + 1) = dMaxX
                    dVerticalMarksY(lVerticalMarks + 1) = (dMinY + dMaxY) / 2
                    lVerticalMarks = lVerticalMarks + 2
                    
                    ReDim Preserve lGeomTypes(lVerticalMarks + 1)
                    lGeomTypes(lVerticalMarks) = eGeomType
                                    
                    If dMinX < dLowestX Then
                        dLowestX = dMinX
                    End If
                    If dMaxX > dHighestX Then
                        dHighestX = dMaxX
                    End If
                    
                    If dMinY < dFLowestY Then
                        dFLowestY = dMinY
                    End If
                    
                    If dMaxY > dFHighestY Then
                        dFHighestY = dMaxY
                    End If
                End If
                
                If ((lGeomFaceId = JXSEC_WEB_LEFT) Or (lGeomFaceId = JXSEC_WEB_RIGHT)) Then
                    If dMinY < dWebLowestY Then
                        dWebLowestY = dMinY
                    End If
                    
                    If dMaxY > dWebHighestY Then
                        dWebHighestY = dMaxY
                    End If
                End If
            End If
        Next i
        
        Dim oNewMfgGeom2d As IJMfgGeom2d
        Set oNewMfgGeom2d = oFirstGeom2d.Clone
        oNewMfgGeom2d.PutMoniker Nothing
        oNewMfgGeom2d.PutEOC NotEOC
        Dim oLine As IJLine
        
        If lFlangeFaceID = JXSEC_TOP Then
            dLowerLeftX = dLowestX
            dLowerLeftY = dWebHighestY
            dLowerRightX = dHighestX
            dLowerRightY = dWebHighestY
            
            dUpperLeftX = dLowestX
            dUpperLeftY = (dWebHighestY + dFlangeThickness)
            dUpperRightX = dHighestX
            dUpperRightY = (dWebHighestY + dFlangeThickness)
            
        Else
            dLowerLeftX = dLowestX
            dLowerLeftY = (dWebLowestY - dFlangeThickness)
            dLowerRightX = dHighestX
            dLowerRightY = (dWebLowestY - dFlangeThickness)
            
            dUpperLeftX = dLowestX
            dUpperLeftY = dWebLowestY
            dUpperRightX = dHighestX
            dUpperRightY = dWebLowestY
        End If
        
        Set oLine = New Line3d
        oLine.DefineBy2Points dLowerLeftX, dLowerLeftY, 0#, dLowerRightX, dLowerRightY, 0#
        
        Dim oCS As New ComplexString3d
        oCS.AddCurve oLine, True
        oNewMfgGeom2d.PutGeometry oCS
        lGeomCount = lGeomCount + 1
        oGeomCol2d.AddGeometry lGeomCount, oNewMfgGeom2d
        
        Set oCS = New ComplexString3d
        Set oLine = New Line3d
        Set oNewMfgGeom2d = oFirstGeom2d.Clone
        oNewMfgGeom2d.PutMoniker Nothing
        oNewMfgGeom2d.PutEOC NotEOC
        
        oLine.DefineBy2Points dLowerRightX, dLowerRightY, 0#, dUpperRightX, dUpperRightY, 0#
        oCS.AddCurve oLine, True
        oNewMfgGeom2d.PutGeometry oCS
        lGeomCount = lGeomCount + 1
        oGeomCol2d.AddGeometry lGeomCount, oNewMfgGeom2d
        
        Set oCS = New ComplexString3d
        Set oLine = New Line3d
        Set oNewMfgGeom2d = oFirstGeom2d.Clone
        oNewMfgGeom2d.PutMoniker Nothing
        oNewMfgGeom2d.PutEOC NotEOC
        
        oLine.DefineBy2Points dUpperRightX, dUpperRightY, 0#, dUpperLeftX, dUpperLeftY, 0#
        oCS.AddCurve oLine, True
        oNewMfgGeom2d.PutGeometry oCS
        lGeomCount = lGeomCount + 1
        oGeomCol2d.AddGeometry lGeomCount, oNewMfgGeom2d
        
        Set oCS = New ComplexString3d
        Set oLine = New Line3d
        Set oNewMfgGeom2d = oFirstGeom2d.Clone
        oNewMfgGeom2d.PutMoniker Nothing
        oNewMfgGeom2d.PutEOC EOC
        
        oLine.DefineBy2Points dUpperLeftX, dUpperLeftY, 0#, dLowerLeftX, dLowerLeftY, 0#
        oCS.AddCurve oLine, True
        oNewMfgGeom2d.PutGeometry oCS
        lGeomCount = lGeomCount + 1
        oGeomCol2d.AddGeometry lGeomCount, oNewMfgGeom2d
        
        For i = 0 To (lVerticalMarks - 1) Step 1
            If ((Abs(dVerticalMarksX(i) - dLowestX) > 0.001) And _
                 (Abs(dVerticalMarksX(i) - dHighestX) > 0.001)) Then
                Dim bDuplicate As Boolean
                bDuplicate = False
                If (i > 0) Then
                    For j = 0 To (i - 1) Step 1
                        If (Abs(dVerticalMarksX(i) - dVerticalMarksX(j)) < 0.001) Then
                            bDuplicate = True
                        End If
                    Next j
                End If
                
                If bDuplicate = False Then
                    Set oCS = New ComplexString3d
                    Set oLine = New Line3d
                    Set oNewMfgGeom2d = oFirstGeom2d.Clone
                    oNewMfgGeom2d.PutMoniker Nothing
                    oNewMfgGeom2d.PutEOC EOC
                    oNewMfgGeom2d.PutGeometrytype STRMFG_FEATURE_LABEL_MARK
                    
                    If ((lGeomTypes(i) = STRMFG_INNER_CONTOUR) Or _
                         ((lFlangeFaceID = JXSEC_TOP) And (dVerticalMarksY(i) > (dFLowestY + dFHighestY) / 2#)) Or _
                         ((lFlangeFaceID = JXSEC_BOTTOM) And (dVerticalMarksY(i) < (dFLowestY + dFHighestY) / 2#))) Then
                        If UpSide = JXSEC_WEB_LEFT Then
                            oNewMfgGeom2d.FaceId = JXSEC_WEB_RIGHT
                        Else
                            oNewMfgGeom2d.FaceId = JXSEC_WEB_LEFT
                        End If
                    End If
                    
                    If (lFlangeFaceID = JXSEC_TOP) Then
                        oLine.DefineBy2Points dVerticalMarksX(i), (dWebHighestY + dFlangeThickness), 0#, dVerticalMarksX(i), dWebHighestY, 0#
                    Else
                        oLine.DefineBy2Points dVerticalMarksX(i), (dWebLowestY - dFlangeThickness), 0#, dVerticalMarksX(i), dWebLowestY, 0#
                    End If
                    
                    oCS.AddCurve oLine, True
                    oNewMfgGeom2d.PutGeometry oCS
                    lGeomCount = lGeomCount + 1
                    oGeomCol2d.AddGeometry lGeomCount, oNewMfgGeom2d
                End If
            End If
        Next i
    End If
        
End Function
Private Function CreateBuiltUpPCMark(ByVal oProfilePart As Object, ByVal lWebFaceID As Long, _
        ByVal lFlangeFaceID As Long, ByVal oGeomCol3d As IJMfgGeomCol3d)
        
    Dim dFlangeWidth As Double, dFlangeThickness As Double
    Dim oGeom2d As IJMfgGeom2d
    Dim oProfilePartSupport As IJProfilePartSupport
    Dim oPartSupport As IJPartSupport
    
    Set oProfilePartSupport = New ProfilePartSupport
    Set oPartSupport = oProfilePartSupport
    Set oPartSupport.Part = oProfilePart
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfilePart
    
    dFlangeWidth = 0#
    dFlangeThickness = 0#
    oProfilePartSupport.GetFlangeWidth dFlangeWidth
    
    If lFlangeFaceID = JXSEC_TOP Then
        oProfilePartSupport.GetThickness TopFlange, dFlangeThickness
    Else
        oProfilePartSupport.GetThickness BottomFlange, dFlangeThickness
    End If
    
    If ((dFlangeWidth < 0.001) Or (dFlangeThickness < 0.001)) Then
        Exit Function
    End If
    
    Dim oUpsideSurface As IUnknown
    Dim oSurfacePort As IJPort
       
    Dim dX As Double, dY As Double, dZ As Double
    Dim oMfgMGHelper As IJMfgMGHelper
    Dim oRangeMin As IJDPosition, oRangeMax As IJDPosition, oUpdatedRangeMin As IJDPosition, oUpdatedRangeMax As IJDPosition
    
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    Dim oGeometryBound As IJGeometryBound
    Set oGeometryBound = New DGeomOpsBound
    
    Set oSurfacePort = oProfileWrapper.GetSurfacePort(lFlangeFaceID)
    Set oUpsideSurface = oSurfacePort.Geometry
    
    If Not TypeOf oUpsideSurface Is IJPlane Then
        Exit Function
    End If
    
    Dim oFlangePlane As IJPlane, oFlangeFinitePlane As IJPlane
    Dim oFlangeInfPlane As New Plane3d
    
    Set oFlangePlane = oUpsideSurface
    oFlangePlane.GetRootPoint dX, dY, dZ
    oFlangeInfPlane.SetRootPoint dX, dY, dZ
    
    oFlangePlane.GetNormal dX, dY, dZ
    oFlangeInfPlane.SetNormal dX, dY, dZ
    
    oMfgMGHelper.GetRange oProfilePart, oRangeMin, oRangeMax
    
    oGeometryBound.CreateFinitePlane Nothing, oFlangeInfPlane, oRangeMin, oRangeMax, Nothing, oFlangeFinitePlane
    
    Dim oWebPlane As IJPlane
    Dim oWebInfPlane As New Plane3d
        
    Set oSurfacePort = oProfileWrapper.GetSurfacePort(lWebFaceID)
    Set oUpsideSurface = oSurfacePort.Geometry
    
    If Not TypeOf oUpsideSurface Is IJPlane Then
        Exit Function
    End If
    
    Set oWebPlane = oUpsideSurface
    oWebPlane.GetRootPoint dX, dY, dZ
    oWebInfPlane.SetRootPoint dX, dY, dZ
    
    oWebPlane.GetNormal dX, dY, dZ
    oWebInfPlane.SetNormal dX, dY, dZ
    
    Dim oWebFinitePlane As IJPlane
    
    oMfgMGHelper.GetRange oUpsideSurface, oRangeMin, oRangeMax
    Dim oOffsetVec As New DVector, oFlangeNormal As New DVector
    
    Set oOffsetVec = oRangeMax.Subtract(oRangeMin)
    
    oFlangePlane.GetNormal dX, dY, dZ
    oFlangeNormal.Set dX, dY, dZ
    
    If (oOffsetVec.Dot(oFlangeNormal) < 0#) Then
        dX = -1 * dX
        dY = -1 * dY
        dZ = -1 * dZ
    End If
    
    oOffsetVec.Set dX * -10#, dY * -10#, dZ * -10#
    
    Set oUpdatedRangeMin = oRangeMin.Offset(oOffsetVec)
    
    oOffsetVec.Set dX * 10#, dY * 10#, dZ * 10#
    Set oUpdatedRangeMax = oRangeMax.Offset(oOffsetVec)
    
    If (Abs(oUpdatedRangeMax.x - oUpdatedRangeMin.x) < 0.001) Then
        oUpdatedRangeMin.x = oUpdatedRangeMin.x - 0.002
        oUpdatedRangeMax.x = oUpdatedRangeMax.x + 0.002
    End If
    
    If (Abs(oUpdatedRangeMax.y - oUpdatedRangeMin.y) < 0.001) Then
        oUpdatedRangeMin.y = oUpdatedRangeMin.y - 0.002
        oUpdatedRangeMax.y = oUpdatedRangeMax.y + 0.002
    End If
    
    If (Abs(oUpdatedRangeMax.z - oUpdatedRangeMin.z) < 0.001) Then
        oUpdatedRangeMin.z = oUpdatedRangeMin.z - 0.002
        oUpdatedRangeMax.z = oUpdatedRangeMax.z + 0.002
    End If
    
    oGeometryBound.CreateFinitePlane Nothing, oWebInfPlane, oUpdatedRangeMin, oUpdatedRangeMax, Nothing, oWebFinitePlane
       
    Dim oInterWire As IJWireBody
    Set oInterWire = m_oMfgRuleHelper.GetCommonGeometry(oWebFinitePlane, oFlangeFinitePlane, False)
    
    Dim oInterCS As IJComplexString
    oMfgMGHelper.WireBodyToComplexString oInterWire, oInterCS
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeom3d As IJMfgGeom3d
    Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
    oGeom3d.PutGeometry oInterCS
    oGeom3d.PutGeometrytype STRMFG_BUILTUP_CONNECTION_MARK
    oGeom3d.FaceId = lFlangeFaceID
    oGeomCol3d.AddGeometry (oGeomCol3d.Getcount + 1), oGeom3d
    
End Function
Private Function UpdateBuiltUpPCMarks(ByVal oProfilePart As Object, ByVal UpSide As Long, ByVal oGeomCol2d As IJMfgGeomCol2d, _
                        ByVal lFlangeFaceID As Long)
                        
    Dim i As Long, lGeomCount As Long, lGeomFaceId As Long, lMarkCount
    Dim oGeom2d As IJMfgGeom2d
    Dim eGeomType As StrMfgGeometryType
    Dim oMarkGeom2ds(2) As IJMfgGeom2d
    Dim oEndPos As New DPosition
    
    lGeomCount = oGeomCol2d.Getcount
    lMarkCount = 0
    
    For i = 1 To lGeomCount
        Set oGeom2d = oGeomCol2d.GetGeometry(i)
        eGeomType = oGeom2d.GetGeometryType
        lGeomFaceId = oGeom2d.FaceId
        
        If (lMarkCount < 2) Then
            If ((eGeomType = STRMFG_BUILTUP_CONNECTION_MARK) And ((lGeomFaceId = lFlangeFaceID))) Then
                Set oMarkGeom2ds(lMarkCount) = oGeom2d
                lMarkCount = lMarkCount + 1
            End If
        End If
    Next i
    
    If lMarkCount <> 2 Then
        Exit Function
    End If
    
    Dim dCrv1StartX  As Double, dCrv1StartY  As Double, dCrv1StartZ  As Double, dCrv1EndX  As Double, dCrv1EndY  As Double, dCrv1EndZ As Double
    Dim dCrv2StartX  As Double, dCrv2StartY  As Double, dCrv2StartZ  As Double, dCrv2EndX  As Double, dCrv2EndY  As Double, dCrv2EndZ As Double
    Dim oPos1 As New DPosition, oPos2 As New DPosition
    Dim dDist1 As Double, dDist2 As Double
    Dim oLine1 As IJLine, oLine2 As IJLine
    
    Set oLine1 = New Line3d
    Set oLine2 = New Line3d
    
    Dim oCurve1 As IJCurve, oCurve2 As IJCurve
    Set oCurve1 = oMarkGeom2ds(0).GetGeometry
    Set oCurve2 = oMarkGeom2ds(1).GetGeometry
    
    oCurve1.EndPoints dCrv1StartX, dCrv1StartY, dCrv1StartZ, dCrv1EndX, dCrv1EndY, dCrv1EndZ
    oCurve2.EndPoints dCrv2StartX, dCrv2StartY, dCrv2StartZ, dCrv2EndX, dCrv2EndY, dCrv2EndZ
    
    oPos1.Set dCrv1StartX, dCrv1StartY, dCrv1StartZ
    oPos2.Set dCrv2StartX, dCrv2StartY, dCrv2StartZ
    dDist1 = oPos1.DistPt(oPos2)
    
    oPos2.Set dCrv2EndX, dCrv2EndY, dCrv2EndZ
    dDist2 = oPos1.DistPt(oPos2)
    
    If dDist1 < dDist2 Then
        oLine1.DefineBy2Points dCrv1StartX, dCrv1StartY, dCrv1StartZ, dCrv2StartX, dCrv2StartY, dCrv2StartZ
        oLine2.DefineBy2Points dCrv1EndX, dCrv1EndY, dCrv1EndZ, dCrv2EndX, dCrv2EndY, dCrv2EndZ
    Else
        oLine1.DefineBy2Points dCrv1StartX, dCrv1StartY, dCrv1StartZ, dCrv2EndX, dCrv2EndY, dCrv2EndZ
        oLine2.DefineBy2Points dCrv1EndX, dCrv1EndY, dCrv1EndZ, dCrv2StartX, dCrv2StartY, dCrv2StartZ
    End If
        
    Dim oNewMfgGeom2d As IJMfgGeom2d
    Set oNewMfgGeom2d = oMarkGeom2ds(0).Clone
    
    Dim oCS As New ComplexString3d
    oCS.AddCurve oLine1, True
    oNewMfgGeom2d.PutGeometry oCS
    lGeomCount = lGeomCount + 1
    oGeomCol2d.AddGeometry lGeomCount, oNewMfgGeom2d
    
    Set oNewMfgGeom2d = oMarkGeom2ds(0).Clone
    
    Set oCS = New ComplexString3d
    oCS.AddCurve oLine2, True
    oNewMfgGeom2d.PutGeometry oCS
    lGeomCount = lGeomCount + 1
    oGeomCol2d.AddGeometry lGeomCount, oNewMfgGeom2d
    
End Function
