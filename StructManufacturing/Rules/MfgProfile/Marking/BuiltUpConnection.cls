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

Private Const MODULE = "MfgProfileMarking.BuiltUpPC"
Private Const IID_IJMfgSystemMarkInfo As Variant = "{CE13B4E9-C6B0-45D6-99F4-0981D870274E}"

' A couple of constants that decide how big and where fitting marks are created
Private Const PI As Double = 3.14159265358979

Private Const dFittingMarkLength As Double = 0.02
Private Const dFittingMarkDistanceFromStartOfWeb As Double = 0.5
    
Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    PrMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    PrMrkHelpers.UnInitialize
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
    
    CreateAPSProfileMarkings STRMFG_BUILTUP_CONNECTION_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    ' return if crossection is not a builtup
    If (IsBuiltUp(Part) = False) Then Exit Function
    
    ' get the web left and web right contours
    ' get the top flange surface
    ' project the web left and web right contours the flange contour
    
    Dim oProfileSupport As IJProfilePartSupport
    Set oProfileSupport = New ProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = oProfileSupport
    Set oPartSupport.Part = Part
    
    Dim eTSide As GSCADSDPartSupport.ThicknessSide
    eTSide = oProfileSupport.ThicknessSideAdjacentToLoadPoint
    If eTSide = SideUnspecified Then
        eTSide = SideA
    End If
    Dim pWebContourColl As Collection, pMonikerColl As Collection
    Dim pFlangeContourColl As Collection, pMonikerColl2 As Collection
    Dim pWebSB As IJSurfaceBody, pFlangeSB As IJSurfaceBody
    Dim oGeom3d As IJMfgGeom3d
    Dim lGeomCount As Long
    
    Dim bMarkFlange As Boolean
    bMarkFlange = MarkFlange(oProfileSupport)
    
    '  get the web surfacebody and the web contour collection
    ' get webright for for SideA and WebLeft for SideB since we are looking underneath
    If (eTSide = SideA) Then
        If (bMarkFlange) Then
            oProfileSupport.GetProfileContours WebLeftFace, pWebSB, pWebContourColl, pMonikerColl
        Else
            oProfileSupport.GetProfileContours TopFlangeBottomFace, pWebSB, pWebContourColl, pMonikerColl
        End If
    Else
        If (bMarkFlange) Then
            oProfileSupport.GetProfileContours WebRightFace, pWebSB, pWebContourColl, pMonikerColl
        Else
            oProfileSupport.GetProfileContours TopFlangeBottomFace, pWebSB, pWebContourColl, pMonikerColl
        End If
    End If
        
    ' only need to use the "flange" surfacebody
    If (bMarkFlange) Then
        oProfileSupport.GetProfileContours TopFlangeBottomFace, pFlangeSB, pFlangeContourColl, pMonikerColl2
    Else
        If (eTSide = SideA) Then
            oProfileSupport.GetProfileContours WebLeftFace, pFlangeSB, pFlangeContourColl, pMonikerColl2
        Else
            oProfileSupport.GetProfileContours WebRightFace, pFlangeSB, pFlangeContourColl, pMonikerColl2
        End If
    End If
    
    Set pMonikerColl2 = Nothing
    Set pMonikerColl = Nothing
    Set pFlangeContourColl = Nothing
    
    Dim oWB As IJWireBody
    Dim pMidPt As IJDPosition, pNewPos As IJDPosition
    Dim oFlangeNorm As IJDVector, oWebNorm As IJDVector
    Dim oCS As IJComplexString
    Dim oMGHelper As New MfgMGHelper
    
    Set oWB = m_oMfgRuleHelper.GetCommonGeometry(pFlangeSB, pWebSB, False)
    If Not oWB Is Nothing Then
        Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oWB)
        If (oFlangeNorm Is Nothing And oWebNorm Is Nothing) Then
            Set pMidPt = m_oMfgRuleHelper.GetMiddlePoint(oWB)
            On Error Resume Next
            If Not (pMidPt Is Nothing) Then
                If (oFlangeNorm Is Nothing) Then
                    oMGHelper.ProjectPointOnSurfaceBody pFlangeSB, pMidPt, pNewPos, oFlangeNorm
                End If
                ' find the web normal for later use
                Set pNewPos = Nothing
                If (oWebNorm Is Nothing) Then
                    oMGHelper.ProjectPointOnSurfaceBody pWebSB, pMidPt, pNewPos, oWebNorm
                End If
                Set pMidPt = Nothing
            End If
            On Error GoTo ErrorHandler
        End If
    End If
    Set oWB = Nothing
    
    ' create two alignment marks (one on the web and one on the flange) to help line the web
    ' to the flange.  Create a marking line on the flange to align the web (webleft or webright)
    
    Dim oMarkingInfo As MarkingInfo

    If Not oCS Is Nothing Then
    
        ' Find the LandingCurve and thickness direction of the web
        Dim oLandingCurve As IJWireBody
        Dim oThicknessDirection As IJDVector
        oProfileSupport.GetProfilePartLandingCurve oLandingCurve, oThicknessDirection, False, eTSide

        ' Find endpoints of the LandingCurve
        Dim oStartPos As IJDPosition
        Dim oEndPos As IJDPosition
        oLandingCurve.GetEndPoints oStartPos, oEndPos

        ' From the start point, get the CrossSection orientation
        Dim oXVector As IJDVector
        Dim oYVector As IJDVector
        Dim oOrigin As IJDPosition
        oProfileSupport.GetOrientation oStartPos, oXVector, oYVector, oOrigin

        Dim oCurveLongest As IJCurve
        Set oCurveLongest = oCS

        Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
        oCurveLongest.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

        ' Find direction vector of the marking line
        Dim oDir As IJDVector
        Set oDir = New DVector
        oDir.Set dEndX - dStartX, dEndY - dStartY, dEndZ - dStartZ

        Dim oCross As IJDVector
        Set oCross = oThicknessDirection.Cross(oYVector)

        Dim dDot As Double
        dDot = oCross.Dot(oDir)

        Dim FlageDir As String
        FlageDir = "R"

        If dDot > 0 Then
            ' Reverse the direction of the marking line
            ' This is done so that the thickness mark can allways get displayed to the left
            Dim oTempCS As IJComplexString
            oMGHelper.ReverseComplexString oCS, oTempCS
            Set oCS = oTempCS
            FlageDir = "L"
        End If

        If (bMarkFlange) Then
            ' At some point in the future, we should decide to pass face id as
            ' either JXSEC_TOP_FLANGE_RIGHT_BOTTOM or JXSEC_TOP_FLANGE_LEFT_BOTTOM,
            ' not hard code JXSEC_TOP.
            Set oGeom3d = CreateGeom3dObject(oCS, STRMFG_BUILTUP_CONNECTION_MARK, JXSEC_TOP, oMarkingInfo)
        Else
            Set oGeom3d = CreateGeom3dObject(oCS, STRMFG_LOCATION_MARK, UpSide, oMarkingInfo)
        End If
        
        '********************************************************************************
        ' Above we make sure the thickness is always on the left side of the marking line.
        ' Set this as property if the system mark.
        oMarkingInfo.FlangeDirection = FlageDir
        
'        Dim objAtt As IJDAttributes
'        Dim oAttrCol As IJDAttributesCol
'        Dim objAttribute As IJDAttribute
'        Set objAtt = oMarkingInfo
'        Set oAttrCol = objAtt.CollectionOfAttributes(IID_IJMfgSystemMarkInfo)
'        Set objAttribute = oAttrCol.Item("FlangeDirection")
'        If Not objAttribute Is Nothing Then
'            objAttribute.Value = FlageDir ' if dDot is greater than 0 and the curve is not reversed above, set to "R"
'        End If
'        Set objAttribute = Nothing
'        Set oAttrCol = Nothing
'        Set objAtt = Nothing
        '********************************************************************************

        If (eTSide = SideA) Then
            If (bMarkFlange) Then
                oMarkingInfo.name = "WebLeft"
'                m_oMfgRuleHelper.ScaleVector oThicknessDirection, -1
                oMarkingInfo.direction = m_oMfgRuleHelper.GetDirection(oThicknessDirection)
                oMarkingInfo.ThicknessDirection = oThicknessDirection
            Else
                oMarkingInfo.name = "TopFlangeBottomFace"
                m_oMfgRuleHelper.ScaleVector oYVector, -1
                oMarkingInfo.direction = m_oMfgRuleHelper.GetDirection(oYVector)
                oMarkingInfo.ThicknessDirection = oYVector
            End If
        Else
            If (bMarkFlange) Then
                oMarkingInfo.name = "WebRight"
                ' Reverse the direction
                'm_oMfgRuleHelper.ScaleVector oThicknessDirection, -1
                oMarkingInfo.direction = m_oMfgRuleHelper.GetDirection(oThicknessDirection)
                oMarkingInfo.ThicknessDirection = oThicknessDirection
            Else
                oMarkingInfo.name = "TopFlangeTopFace"
                ' Reverse the direction
                m_oMfgRuleHelper.ScaleVector oYVector, -1
                oMarkingInfo.direction = m_oMfgRuleHelper.GetDirection(oYVector)
                oMarkingInfo.ThicknessDirection = oYVector
            End If
        End If
        
        oMarkingInfo.FittingAngle = PI / 2

        oGeomCol3d.AddGeometry 1, oGeom3d
        Set oGeom3d = Nothing
        Set oMarkingInfo = Nothing
        ' create 2 fitting marks to align the flange and the web

        ' If you want the fitting mark to appear in the overlapping portion of flange/web,
        ' leave the below two lines commented.  Uncommenting the two lines below causes
        ' fitting marks to be placed 500 mm from the start point of the Part (there may or
        ' may not be material for this location!) -- Anand for TR 112975
        '
        ' m_oMfgRuleHelper.ExtendWire oCS, 0.5 * oCurveLongest.Length
        ' TrimToActualLength oCS, Part, UpSide

        Dim oTempWireBody As IJWireBody
        Set oTempWireBody = m_oMfgRuleHelper.ComplexStringToWireBody(oCS)

        ' What if the length of the overlapping portion is smaller than (or almost the same as)
        ' the value of "dFittingMarkDistanceFromStartOfWeb" constant (default value is 500mm)?
        ' Set it to ONE-THIRD of the overlapping portion's length.
        
        Dim dEndFittingMarkDist As Double
        Dim dStartFittingMarkDist As Double
        dStartFittingMarkDist = oCurveLongest.Length
        If dStartFittingMarkDist < 1.5 * dFittingMarkDistanceFromStartOfWeb Then
            dStartFittingMarkDist = dStartFittingMarkDist / 3
        Else
            dStartFittingMarkDist = dFittingMarkDistanceFromStartOfWeb
        End If
        
        Dim oFittingCS As IJComplexString
        Set oFittingCS = CreateFittingCurve(oTempWireBody, dStartFittingMarkDist, oFlangeNorm, dFittingMarkLength)
        
        If Not oFittingCS Is Nothing Then
            If (bMarkFlange) Then
                Set oGeom3d = CreateGeom3dObject(oFittingCS, STRMFG_FITTING_MARK, UpSide, oMarkingInfo)
                oMarkingInfo.name = "FlangeFitting"
            Else
                ' At some point in the future, we should decide to pass face id as
                ' either JXSEC_TOP_FLANGE_RIGHT_BOTTOM or JXSEC_TOP_FLANGE_LEFT_BOTTOM,
                ' not hard code JXSEC_TOP.
                
                Set oGeom3d = CreateGeom3dObject(oFittingCS, STRMFG_FITTING_MARK, JXSEC_TOP, oMarkingInfo)
                oMarkingInfo.name = "WebFitting"
            End If
            
            oGeomCol3d.AddGeometry 1, oGeom3d
            Set oGeom3d = Nothing
            Set oMarkingInfo = Nothing
            Set oFittingCS = Nothing
        End If
        
        ' Create another fitting mark with an offset from the end position
        dEndFittingMarkDist = oCurveLongest.Length - dStartFittingMarkDist
        Set oFittingCS = CreateFittingCurve(oTempWireBody, dEndFittingMarkDist, oFlangeNorm, dFittingMarkLength)
        
        If Not (oFittingCS Is Nothing) Then
            If (bMarkFlange) Then
                Set oGeom3d = CreateGeom3dObject(oFittingCS, STRMFG_FITTING_MARK, UpSide, oMarkingInfo)
                oMarkingInfo.name = "FlangeFitting"
            Else
                ' At some point in the future, we should decide to pass face id as
                ' either JXSEC_TOP_FLANGE_RIGHT_BOTTOM or JXSEC_TOP_FLANGE_LEFT_BOTTOM,
                ' not hard code JXSEC_TOP.
                
                Set oGeom3d = CreateGeom3dObject(oFittingCS, STRMFG_FITTING_MARK, JXSEC_TOP, oMarkingInfo)
                oMarkingInfo.name = "WebFitting"
            End If
            
            oGeomCol3d.AddGeometry 1, oGeom3d
            Set oGeom3d = Nothing
            Set oMarkingInfo = Nothing
            Set oFittingCS = Nothing
        End If

        ' make the fitting mark on the flange long enough so it doesn't matter which side the "FlangeFitting"
        ' mark is on the web
        If Not (oWebNorm Is Nothing) Then
            Set oFittingCS = CreateFittingCurve(oTempWireBody, dStartFittingMarkDist, oWebNorm, dFittingMarkLength)
            If Not (oFittingCS Is Nothing) Then
                If (bMarkFlange) Then
                    ' At some point in the future, we should decide to pass face id as
                    ' either JXSEC_TOP_FLANGE_RIGHT_BOTTOM or JXSEC_TOP_FLANGE_LEFT_BOTTOM,
                    ' not hard code JXSEC_TOP.
                    
                    Set oGeom3d = CreateGeom3dObject(oFittingCS, STRMFG_FITTING_MARK, JXSEC_TOP, oMarkingInfo)
                    oMarkingInfo.name = "WebFitting"
                Else
                    Set oGeom3d = CreateGeom3dObject(oFittingCS, STRMFG_FITTING_MARK, UpSide, oMarkingInfo)
                    oMarkingInfo.name = "FlangeFitting"
                End If
                
                oGeomCol3d.AddGeometry 1, oGeom3d
                Set oGeom3d = Nothing
                Set oMarkingInfo = Nothing
                Set oFittingCS = Nothing
            End If
            
            ' Create another fitting mark with an offset from the end position
            Set oFittingCS = CreateFittingCurve(oTempWireBody, dEndFittingMarkDist, oWebNorm, dFittingMarkLength)
            If Not (oFittingCS Is Nothing) Then
                If (bMarkFlange) Then
                    ' At some point in the future, we should decide to pass face id as
                    ' either JXSEC_TOP_FLANGE_RIGHT_BOTTOM or JXSEC_TOP_FLANGE_LEFT_BOTTOM,
                    ' not hard code JXSEC_TOP.
                    
                    Set oGeom3d = CreateGeom3dObject(oFittingCS, STRMFG_FITTING_MARK, JXSEC_TOP, oMarkingInfo)
                    oMarkingInfo.name = "WebFitting"
                Else
                    Set oGeom3d = CreateGeom3dObject(oFittingCS, STRMFG_FITTING_MARK, UpSide, oMarkingInfo)
                    oMarkingInfo.name = "FlangeFitting"
                End If
                
                oGeomCol3d.AddGeometry 1, oGeom3d
                Set oGeom3d = Nothing
                Set oMarkingInfo = Nothing
                Set oFittingCS = Nothing
            End If
            
        End If
    End If

    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
CleanUp:
    Set oGeomCol3d = Nothing
    Set pWebContourColl = Nothing
    Set pWebSB = Nothing
    Set pFlangeSB = Nothing
    Set oMGHelper = Nothing
    Set oCS = Nothing
    Set oMarkingInfo = Nothing
    Set oFittingCS = Nothing
    Set oWebNorm = Nothing
    Set oFlangeNorm = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2002, , "RULES")
    GoTo CleanUp
End Function

Private Sub TrimToActualLength(oCS As IJComplexString, ByVal Part As Object, UpSide As Long)
    Const METHOD = "TrimToActualLength"
    On Error GoTo ErrorHandler
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    ' Temporary 3D collection
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)

    ' Create temporary origin marks
    Dim oMfgHelper As IJMfgEntityHelper
    Set oMfgHelper = New MfgEntityHelper
    oMfgHelper.CreateOriginMarks Part, UpSide, oGeomCol3d

    Dim oPositions As New Collection
    
    ' Extract all the possible candidate positions
    Dim h As Long
    For h = 1 To oGeomCol3d.Getcount
    
        Dim oTestGeom3d As IJMfgGeom3d
        Set oTestGeom3d = oGeomCol3d.GetGeometry(h)
        If oTestGeom3d.GetGeometryType = STRMFG_PROFILE_ORIGIN Then
        
            Dim oGeom3dCS As IJComplexString
            Set oGeom3dCS = oTestGeom3d.GetGeometry
            
            Dim oCurve As Object
            oGeom3dCS.GetCurve 1, oCurve
            If oGeom3dCS.CurveCount = 1 And TypeOf oCurve Is IJLineString Then
                Dim oLineString As IJLineString
                Set oLineString = oCurve
                
                Dim lPointCount As Long
                lPointCount = oLineString.PointCount
                
                Dim i As Long
                For i = 1 To lPointCount
                    Dim dX As Double, dY As Double, dZ As Double
                    oLineString.GetPoint i, dX, dY, dZ
                   
                    Dim oPosR As IJDPosition
                    Set oPosR = New DPosition
                    oPosR.Set dX, dY, dZ
                    oPositions.Add oPosR
                Next
            Else
                Dim JCurve As IJCurve
                Set JCurve = oGeom3dCS
                Dim StartX As Double, StartY As Double, StartZ As Double, endX As Double, endY As Double, endZ As Double
                JCurve.EndPoints StartX, StartY, StartZ, endX, endY, endZ
                
                Dim oPos1 As IJDPosition
                Set oPos1 = New DPosition
                oPos1.Set StartX, StartY, StartZ
                
                Dim oPos2 As IJDPosition
                Set oPos2 = New DPosition
                oPos2.Set endX, endY, endZ
                
                oPositions.Add oPos1
                oPositions.Add oPos2
            End If
        End If
    Next
    
    If oPositions.Count < 1 Then
        ' We found no points (this should not happen, but for the sake of survivability we just continue)
        Exit Sub
    End If

    ' Convert to WireBody
    Dim oWireBody As IJWireBody
    Set oWireBody = m_oMfgRuleHelper.ComplexStringToWireBody(oCS)

    ' Project all points down to the extended curve
    Dim oProjectedPositions As New Collection
    Dim oPosition As IJDPosition
    For Each oPosition In oPositions
        Dim oProjectedPosition As IJDPosition
        Set oProjectedPosition = m_oMfgRuleHelper.ProjectPointOnWire(oWireBody, oPosition)
        oProjectedPositions.Add oProjectedPosition
    Next

    If oProjectedPositions.Count < 1 Then
        ' We found no points (this should not happen, but for the sake of survivability we just continue)
        Exit Sub
    End If

    ' Select the 2 points that are the farthest apart
    Dim dLongestDist As Double
    dLongestDist = 0
    
    Dim oProjLongestPos1 As IJDPosition
    Dim oProjLongestPos2 As IJDPosition
    Dim oProjPos1 As IJDPosition
    Dim oProjPos2 As IJDPosition
    For Each oProjPos1 In oProjectedPositions
        For Each oProjPos2 In oProjectedPositions

            Dim dTempDist As Double
            dTempDist = oProjPos1.DistPt(oProjPos2)
            If dTempDist > dLongestDist Then
                dLongestDist = dTempDist
                Set oProjLongestPos1 = oProjPos1
                Set oProjLongestPos2 = oProjPos2
            End If
        Next
    Next

    Dim oMG As IJMfgMGHelper
    Set oMG = New MfgMGHelper

    ' Do the actual trimming
    oMG.TrimCurveByPoints oCS, oProjLongestPos1, oProjLongestPos2
    
    'Cleanup
    Dim oDObject As IJDObject
    Set oDObject = oGeomCol3d
    oDObject.Remove
    
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2002, , "RULES")
End Sub

'*******************************************************************
'   CreateFittingCurve
'      Create a line from the inputs.  Find a baseline position on oCS given the dParm value.
'      Split the length in half and create the Start and end points of the line by offsetting the baseline
'      point along the direction vector by half the input length
'*******************************************************************

Private Function CreateFittingCurve(oWireBody As IJWireBody, dDistanceFromStart As Double, oDirection As IJDVector, dLengthOfCurve As Double) As IJComplexString
Const METHOD = "CreateFittingCurve"
On Error GoTo ErrorHandler
    
    Dim oLine As New Line3d
    Dim x As Double, y As Double, z As Double, xDiff As Double, yDiff As Double, zDiff As Double
    
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    oWireBody.GetEndPoints oStartPos, oEndPos

    Dim oRefPos As IJDPosition
    Set oRefPos = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oWireBody, oStartPos, dDistanceFromStart, oEndPos)

    x = oRefPos.x
    y = oRefPos.y
    z = oRefPos.z
    
    xDiff = (dLengthOfCurve * oDirection.x) / 2
    yDiff = (dLengthOfCurve * oDirection.y) / 2
    zDiff = (dLengthOfCurve * oDirection.z) / 2

    ' split the difference for the length of the curve
    oLine.DefineBy2Points (x - xDiff), (y - yDiff), (z - zDiff), (x + xDiff), (y + yDiff), (z + zDiff)
    Dim oFittingCS As New ComplexString3d
    oFittingCS.AddCurve oLine, True

    Set CreateFittingCurve = oFittingCS

    Set oLine = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2002, , "RULES")
End Function

'*******************************************************************
'   ModifyCurve - shorten the input complexstring
'      Trim the input complexstring by two points.  The points are defined by the input parameters
'
'*******************************************************************

Private Sub ModifyCurve(oCS As IJComplexString, dStartParm As Double, dEndParm As Double)
Const METHOD = "ModifyCurve"
On Error GoTo ErrorHandler
    
    Dim x As Double, y As Double, z As Double
    Dim oNewStartPos As New DPosition
    Dim oNewEndPos As New DPosition
    Dim oCurve As IJCurve
    Set oCurve = oCS
    oCurve.Position dStartParm, x, y, z
    oNewStartPos.Set x, y, z
    oCurve.Position dEndParm, x, y, z
    oNewEndPos.Set x, y, z
    
    Dim oMGHelper As New MfgMGHelper
    oMGHelper.TrimCurveByPoints oCS, oNewStartPos, oNewEndPos
    
    Set oCurve = Nothing
    Set oMGHelper = Nothing
    Set oNewStartPos = Nothing
    Set oNewEndPos = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2002, , "RULES")
'    ReportUnanticipatedError "Helper", "CreateFittingCurve"
End Sub


'*******************************************************************
'   MarkFlange
'      Determine if the flange or web needs to be marked
'*******************************************************************

Private Function MarkFlange(oProfileSupport As IJProfilePartSupport) As Boolean
Const METHOD = "MarkFlange"
On Error GoTo ErrorHandler
    
    MarkFlange = True
    Dim oXSect As IJCrossSection
    oProfileSupport.GetCrossSection oXSect
    Dim oPC As IJDPartClass
    Set oPC = oXSect.GetPartClass
    Dim oXSectPC As IJDCrossSectionPartClass
    Set oXSectPC = oPC
        
    Dim oXSectType As IJDCrossSectionType
    Set oXSectType = oXSectPC.CrossSectionType
    
    Dim sName As String
    sName = oXSectType.name
    
    If Not (LCase(sName) = LCase("Teebar")) Then MarkFlange = False

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2002, , "RULES")
    'ReportUnanticipatedError "Helper", "CreateFittingCurve"
End Function


