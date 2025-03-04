VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserDefinedValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------------
'     Copyright (C) 1998-2004 Intergraph Corporation. All rights reserved.
'
'
' File Info:
'     Folder:  S:\StructManufacturing\Data\MfgRules\MfgProfileProcess
'     Project: MfgProfileProcess
'     Module:  UserDefinedValues.cls
'
' Abstract:
'     Provides methods for setting parameters in the unfold algorithm.
'
' History:
'     Christian List    2004.01.02 Created.
'
'-----------------------------------------------------------------------------------------
Option Explicit
Const MODULE = "MfgProfileProcess"

Implements IJDMfgUnfoldParameters

Private Const sSOURCEFILE As String = "UserDefinedValues.cls"
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors   ' To collect and propagate the errors.
Private m_oError As IJEditError     ' Defined here for convenience

Private m_oObject As Object

Private Sub Class_Initialize()

    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject(ERRORPROGID)
    End If
    
    Set m_oObject = Nothing
    
End Sub

Private Sub Class_Terminate()

    Set m_oObject = Nothing
    Set m_oErrors = Nothing

End Sub

Private Function IJDMfgUnfoldParameters_GetDoubleValue(ByVal bstrParameterName As String) As Double
    Const METHOD As String = "IJDMfgUnfoldParameters_GetDoubleValue"
    On Error GoTo ErrorHandler
    
    Select Case bstrParameterName
    Case "DistBendLine"
        'Distance between calculated bending lines.
        'The smaller this value is, the more accurate the unfolding will be,
        'and the more calculation time is required.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.25

    Case "DistEvalCurvature"
        'Distance between calculated bending lines for evaluating curvature.
        'The smaller this value is, the more accurate the evaluation will be,
        'and the more calculation time is required.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.5

    Case "MaxCurvature"
        'A bending line is calculated as the intersection between the plate and a plane.
        'If the intersection curve has a curvature less than MaxCurvature,
        'the intersection curve can be used as a bending line.
        'The curvature is calculated as the chord height for the whole intersection curve divided by the curve length.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.002

    Case "AngleTolerance"
        'The intersection curve with minimum curvature between plate and plane is calculated by iteration.
        'When the plane is positioned within an accuracy less than AngleTolerance, the iteration stops.
        'AngleTolerance is defined in degrees.
        'The lesser the value of AngleTolerance, the higher will be the accuracy with which the bending lines are calculated.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.01

    Case "PointsOnLine"
        'After unfolding, curves are converted from points into lines.
        'If the distance from a line through start and end point to all points is less than this value,
        'the points are substituted by a line.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.001

    Case "LongProcessLine"
        'The bending line through the center of the plate is used for the Processing Line.
        'If another bending line is at least LongProcessLine longer,
        'this bending line becomes the Processing Line.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.25

    Case "Zmax"
        'A plate is plane, if all points on the plate are with a distance of this value from a plane,
        'defined through 3 points on the plate.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.002

    Case "DistPointLine"
        'Distance from a point to a line.
        'If the distance is less than this value, the point is considered to be on the line.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0001

    Case "DL"
        'When the knuckle lines are converted into bending lines, the knuckle lines are made longer.
        'This is done to avoid numerical problems with the outer contour.
        'The start and end points of a knuckle line are moved DL outside the plate.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.01

    Case "PointDist"
        'To unfold the curves on the plate, the curves are converted into points.
        'The distance between these points is defined by this value.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.1

    Case "EpsRollBoundary"
        'Tolerance for finding RollBoundaryLines.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0001

    Case "ChordHeight"
        'When the profile is divided into segments,
        'the landing curve is first divided into segments, where the chord height is max this value.
        'The minimum segment length defined by ChordHeigth and SegmentLength is used
        IJDMfgUnfoldParameters_GetDoubleValue = 0.001
  
    Case "SegmentLength"
        'When the profile is divided into segments, the landing curve is first divided into segments,
        'where the segments is max this length.
        'The minimum segment length defined by ChordHeigth and SegmentLength is used
        IJDMfgUnfoldParameters_GetDoubleValue = 0.1
  
    Case "Overlap"
        'When defining inverse bending lines, and more than one line is needed,
        'there is an overlap of these lines, defined by this value
        IJDMfgUnfoldParameters_GetDoubleValue = 0.8

    Case "SegmentFactor"
        'For evaluating curvature of a profile,
        'in may not be necessary to divide the profile into the same number of segments as for profile unfold.
        'The values SegmentLength and ChordHeight are multiplied by this value when evaluating profile curvature.
        IJDMfgUnfoldParameters_GetDoubleValue = 2

    Case "LimitNoBend"
        'When the profile curvature in the profile XY-plane (plane of web) is less than this value,
        'the profile is considered to have no bend.
        'Curvature is the max distance from a line through the endpoints of the profile location line to the location line (curve).
        'LimitNoBend is max curvature divided by profile length.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.05

    Case "LimitNoRoll"
        'When the profile curvature in the profile XZ-plane (plane of flange) is less than this value,
        'the profile is considered to have no roll.
        'Curvature is the max distance from a line through the endpoints of the profile location line to the location line (curve).
        'LimitNoRoll is max curvature divided by profile length.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.05

    Case "LimitNoTwist"
        'When the max twist of a profile is less than this value,
        'the profile is considered to have no twist.
        'LimitNoTwist is given in radians.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.157
        
    Case "EndCutTolerance"
        'The amount to shrink profile contours, to allow for endcuts to go all
        'the way through the contour.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.00001

    Case "LimitNoBendInverseBendingLines"
        'When the profile curvature in the profile XY-plane (plane of web) is less
        'than this value, the profile is considered to not need inverse bending lines.
        'Curvature is the max distance from a line through the endpoints of the
        'profile location line to the location line (curve).
        'LimitNoBendInverseBendingLines is max curvature divided by profile length
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0005

    Case "LimitNoRollInverseBendingLines"
        'When the profile curvature in the profile XZ-plane (plane of flange) is less
        'than this value, the profile is considered to not need inverse bending lines.
        'Curvature is the max distance from a line through the endpoints of the
        'profile location line to the location line (curve).
        'LimitNoRollInverseBendingLines is max curvature divided by profile length
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0005
        
    Case "MinimumCurveLengthForNesting"
        'Minimum Curve Length in the final output before going to nesting
        IJDMfgUnfoldParameters_GetDoubleValue = 0.001
        
    Case "ChordHeightToleranceForArcConversion"
        'An arc is converted to a line, if the chord height is less than this value
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0001
        
    Case "LimitIBLStraightRatio"
        ' The Profile is first unbend into XY plane.
        ' The straight portion of the Profile does not require Inverse bending lines
        ' if ratio straight portion w.r.t total length is greater than "LimitIBLStraightRatio"
        IJDMfgUnfoldParameters_GetDoubleValue = 0.5

    Case "MaximumFeatureLengthForMerging"
        ' Sometimes the profile contours created by feature end cuts are too small to apply any bevels.
        ' If this MaximumFeatureLengthForMerging parameter value is defined and if it is more than 0.001 mm
        ' Then all the end cut curves smaller than this values will be merged with the adjacent bigger segments.
        '  Curves will be merged only if both adjacent segments are lines and they are collinear.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.015
        
    Case "MinimumFeatureSegmentLength"
        ' After imposing feature cuts, any resulting contour curves less than this value, will be removed from the group.
        ' Adjacent contours will be extended and trimmed to close the gap.
        IJDMfgUnfoldParameters_GetDoubleValue = 0#
        
    Case "ChordHeightToleranceForStrokingBsplineCurve"
        ' Given a B-spline curve and approximation tolerance, we generate number of points on the curve
        ' such that the line created by two adjacent points approximates the B-spline curve locally
        ' within this tolerance (chord-height < tolerance)
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0001
        
    Case "TopFlangeLength"
        ' Get The Top Flange length from the profile part.
        ' This will be the width of the top flange shape in part monitor output.
        ' If user choose not to implement this property, TopFlange Length will be added to FilletRadius as the default output.
        If TypeOf m_oObject Is IJProfilePart Then
            Dim oProfilePartSupport As IJProfilePartSupport
            Dim oPartSupport As IJPartSupport
            
            Set oProfilePartSupport = New ProfilePartSupport
            Set oPartSupport = oProfilePartSupport
            Set oPartSupport.Part = m_oObject
            
            oProfilePartSupport.GetFlangeWidth IJDMfgUnfoldParameters_GetDoubleValue
            Set oPartSupport = Nothing
            Set oProfilePartSupport = Nothing
        Else
            IJDMfgUnfoldParameters_GetDoubleValue = 0#
        End If
        
     Case "MaxOverlapDistForMarkingAndOuterContour"
        ' If the maximum distance between a marking line and the outer contour
        ' is less than this value, that marking line will be deleted.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.001
               
    Case "ProfileFaceOffset"
        ' If this value is set to a value greater than 0 the individual profile faces will
        ' be shown in the Part Monitor with a gap between the faces
        ' corresponding to this value
        'IJDMfgUnfoldParameters_GetDoubleValue = 0.05 ' 50mm gap between the faces
        IJDMfgUnfoldParameters_GetDoubleValue = -1 ' No Offset between the faces
                
     Case "PartRotationAngle"
        ' If the profile XML data needs to be rotated by certain angle, user specifies the amount of degrees with this parameter
        ' Only positive values are expected. Zero or negative values will be ignored by the system.
                IJDMfgUnfoldParameters_GetDoubleValue = 0# ' Default - No Rotation
        'IJDMfgUnfoldParameters_GetDoubleValue = GetPartRotationAngleForER
                
    Case Else
        IJDMfgUnfoldParameters_GetDoubleValue = -1
        
    End Select

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2034, , "RULES")
End Function

Private Function IJDMfgUnfoldParameters_GetLongValue(ByVal bstrParameterName As String) As Long
    Const METHOD As String = "IJDMfgUnfoldParameters_GetLongValue"
    On Error GoTo ErrorHandler
    
    Select Case bstrParameterName
    Case "NumPoints"
        'Number of points calculated on intersection curve between plane and surface.
        'Used for finding bending lines (curvature of plate)
        IJDMfgUnfoldParameters_GetLongValue = 20
                
    Case "MarginsShrinkagesOrder"
        ' Margin and shrinkages are applied on manufacturing part in post-unfolding stage.
        ' With this parameter, user can control, in which order margin and shrinkages are applied
        ' 0 - Random Order
        ' 1 - Apply Margins, ShrinkageMargins and Shrinkages in that order
        ' 2 - Apply ShrinkageMargins, Margins and Shrinkages in that order
        ' 3 - Apply Shrinkages, Margins and ShrinkageMargins in that order
        ' 4 - Apply Shrinkages, ShrinkageMargins and Margins in that order
        IJDMfgUnfoldParameters_GetLongValue = 1
                
    Case "NFlatTest"
        'Number of points tested in U- and V- direction when testing plate for being plane.
        'NFlatTest*NFlatTest points are tested.
        IJDMfgUnfoldParameters_GetLongValue = 20
        
            ' add new cases
            Case "InCurvature"
                        IJDMfgUnfoldParameters_GetLongValue = StartAtBottom
            Case "OutCurvature"
                        IJDMfgUnfoldParameters_GetLongValue = StartAtTop
            Case "SCurved"
                        IJDMfgUnfoldParameters_GetLongValue = StartAtNeutralAxis ' or maybe StartAtBottom
            Case "Straight"
                        IJDMfgUnfoldParameters_GetLongValue = StartAtNeutralAxis
        
    Case "AverageMethodForVaryingBevelAngle"
        ' When outer contour of a manufactured plate has a varying bevel against another bounding
        ' entity, the manufactured plate output splits the outer contour at specific places as
        ' defined by IJDMfgPlateBevelRule_GetBevelDeviationAngle.
        ' The splits are done such that the start and end points of a region have angles that are
        ' AT MOST 'IJDMfgPlateBevelRule_GetBevelDeviationAngle' apart.
        
        ' The property here, controls how the average angle WITHIN each region is calculated.
        ' 0  System default (decided by Intergraph)
        ' 1  Half of (Attachment angle at region-start + Attachment angle at region-end)
        ' 2  Weighted average of all attachment angles calculated at points within this region.
        ' 3  Smallest of all attachment angles calculated at points within this region.
        ' 4  Largest of all attachment angles calculated at points within this region.
        ' 5  Of all attachment angles calculated in region, the one farthest from 90 degrees.
        ' 6  Of all attachment angles calculated in region, the one closest to 90 degrees.
                                                
        IJDMfgUnfoldParameters_GetLongValue = 1 ' Half of (Att angle at Start + Att angle at end)

    Case "MergeChildFeatureContours"
        ' If a struct feature is child of another parent feature, in the final XML,
        ' both the parent edge and child edges will be merged as a single edge, based on this flag.
        ' 0 - Do not merge. Each feature edge will remain separate SMS_EDGE.
        ' 1 - Merge parent and child feature edges into one single SMS_EDGE.
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "ApplyScalingOnFeatures"
       'With this parameter, user has ability to control how scaling is applied on features ( Sketched features, slots, scallops, corner features, etc.. ).
       'If the flag is set to 0, we do not apply scaling to features at all ( current behavior ).
       'If the flag is set to 1, we always apply scaling to features.
       'If the flag is set to 2, we apply scaling to features for normal shrinkage, but No scaling of features in scaling margin case.
       'If the flag is set to 3, feature scaling is done for scaling margin, but not for regular shrinkage.
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "LandingCurveDefinition"
        ' Profile part's landing curve is the most accurate representation of a profile,
        ' but its geometry does not render easily to roll boundary identification.
        
        ' The molded form definition (e.g., say if sketched) preserves the lines and
        ' arcs as-is, so roll boundaries may be identified.  However, it may not correspond
        ' to the actual profile part's location.
        
        ' We try to do an offset and a trim of the molded form defintion to the
        ' profile part location to get the best of both.  However, the offset routine
        ' is error prone, and the trim routine's logic could go wrong for landing curves
        ' that form closed curves.
        
        ' This option allows the user to specify which of the two alternatives is to be picked.
        ' ** 1 == Use Profile Part landing curve.
        ' ** 2 == Get Molded form definition, offset and trim to part location.
        
        ' Set Profile part landing curve as the (conservative) default.
        IJDMfgUnfoldParameters_GetLongValue = 1

        ' Override for cases where the new molded form based landing curve definition will work.
        ' Instances are Edge reinforcements or when stiffened plate is planar.

        If TypeOf m_oObject Is IJProfileER Then
            IJDMfgUnfoldParameters_GetLongValue = 2
        ElseIf TypeOf m_oObject Is IJStiffenerPart Then
            Dim oStiffPart As IJStiffenerPart
            Set oStiffPart = m_oObject

            On Error Resume Next
            Dim MountingFacePort As IJPort
            oStiffPart.GetMountingFacePort MountingFacePort, ""
            Err.Clear

            If TypeOf MountingFacePort.Geometry Is IJPlane Then
                IJDMfgUnfoldParameters_GetLongValue = 2
            End If
            Err.Clear
            On Error GoTo ErrorHandler

            Set oStiffPart = Nothing
            Set MountingFacePort = Nothing
        End If

    Case "LineArcApproximationAlgorithm"
        ' While getting XML data for part monitor command, each curve in the manufacturing output goes
        ' through a line/arc approximation algorithm to convert the curves to lines and arcs.
        ' With this process parameter, user can control which line/arc approximate to use.
        ' 1 - GTYPE Line/Arc Approximation Algorithm implemented by using GTYPE curves
        ' 2 - ACIS Line Arc Approximation Algorithm implemented by using ACIS Edges
        IJDMfgUnfoldParameters_GetLongValue = 1
        
    Case "AccumulateShrinkages"
        ' If there are multiple shrinkages on a plate part, this parameter controls how to accumulate them.
        ' This is a number which is pointing to an SRD rule
        ' 151 - Add shrinkages
        ' 152 - Multiply Shrinkages
                ' 153 - Average of Shrinkages
        IJDMfgUnfoldParameters_GetLongValue = 151
        
    Case "ApproximateFeatureSplineWithLineOrArc"
        ' With this parameter, user has ability to change the bspline curves in a feature objects to a line or arc.
        ' 0 - Keeep the bspline curves as is
        ' 1 - Convert bspline curves to arc, if possible.
        ' 2 - Convert bspline curves to line, if possible.
        ' 3 - Convert bspline curves to line or arc, if possible.
        IJDMfgUnfoldParameters_GetLongValue = 0
            
    Case "FeatureMoveDirection"
        ' This flag controls how to move the feature curves while applying margin.
        ' 0 - Move feature geometry normal to the margin edge.
        ' 1 - Move feature geometry along the bounding non-feature edge
        IJDMfgUnfoldParameters_GetLongValue = 1
        
    Case "ApplyGeometryBevelOnEdges"
        ' If a feature doesn't have any PC/FET associated with it, this parameter will let the system compute long-point bevel automatically.
        '    0 - Do Not apply any Geometry based varying bevel on edge
        '    1 - Apply Geometry based varying bevel on Feature Edges
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "MacroLibraryIdentifier"
        ' This flag specifies the identifier for the library of macro definitions.
        ' This identifier is used to locate the record in the JMfgMacroDefOption catalog view.
        ' OOTB rules follow the convention that Detail Profile Parts have this ID as 100
        IJDMfgUnfoldParameters_GetLongValue = 100
        
    Case "MfgProfileOrigin"
        ' This specifies what is to be taken as the reference for X = 0
        ' for all the graphics in the Mfg output XML for the Mfg Profile
        ' 0 - (Default) Use the left most Profile Origin point (defined in the end-cut symbol)
        ' 1 - Use the material extremity point on the left (on the outer contour)
        ' 2 - Use the WEB face's material extremity point on the left (on the outer contour)
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "FeatureOrientationProcess"
        ' This specifies the methodology/process reg. how the cutting tool of
        ' feature-cuts are oriented, post-unfold.
        ' 0 - Legacy. Tiny coordinate axes are created as marks for feature
        '     locations. These feature marks are unfolded, and the respective
        '     cutting tools are oriented based on the orientation of the
        '     unfolded coordinate axes.
        ' 1 - Contour-Aligned. Marks corresponding to edges due to feature
        '     cuts, are are created in 3D. These marks are unfolded (and
        '     could get deformed, esp for very curved profiles). The cutting
        '     tool is oriented to align with these unfolded marks to the
        '     extent possible.
        IJDMfgUnfoldParameters_GetLongValue = 1
		
    Case "CloseOptionForFillingMarginGaps"
        ' Applying margin to an edge leave gaps between edge geometry and the
        ' adjacent edge geometry.  This option tells the system how to connect
        ' the gap
        ' 1 - Edge geometry follows the natural/polynomial extension
        ' 2 - Gaps will be connected a new line segment
        ' 3 - Edge geometry follows the tangential extension
        ' 4 - (NB: Applies only to gap created by margin that moves the
        '          Web-Cut edge and fixed circular-arc shaped corner feature)
        '     Extend OR Retract sweep angle of circular-arc edge's sweep angle
        '     to 90 degrees, then tangentially extend both edges until they
        '     intersect.
        IJDMfgUnfoldParameters_GetLongValue = 1
        
    Case "ExtractProfileSketchGeometries"
        ' With this option user can control the ability to extract profile sketch curves and dimensions into Mfg part xml
        ' 0 - Do not extract profile sketch data
        ' 1 - Extract profile sketch data
        ' Invalid Option - same as 0
        IJDMfgUnfoldParameters_GetLongValue = 0
                     
    Case Else
        IJDMfgUnfoldParameters_GetLongValue = -1
        
    End Select

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2035, , "RULES")
End Function

Private Sub IJDMfgUnfoldParameters_GetKnuckleParameters(ByVal pKnuckle As Object, pdRadius As Double, plBendOrRolled As Long)
    Const METHOD As String = "IJDMfgUnfoldParameters_GetKnuckleParameters"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Property Let IJDMfgUnfoldParameters_Object(ByVal RHS As Object)
    Const METHOD As String = "IJDMfgUnfoldParameters_Object"
    On Error GoTo ErrorHandler
    
    Set m_oObject = RHS

    Exit Property
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property
  Private Function GetPartRotationAngleForER() As Double
    GetPartRotationAngleForER = 0#
    
    ' Check if this is a profile ER.
    If TypeOf m_oObject Is IJProfileER Then
        Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
        Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
        Set oProfileWrapper.object = m_oObject
        
        ' Get MfgProfile of the current profile
        Dim oMfgProfilePart As IJMfgProfilePart
        oProfileWrapper.ProfileHasMfgPart oMfgProfilePart
        
        ' Get the stiffened Plate
        Dim oStiffenedPlates As IJElements
        Set oStiffenedPlates = oProfileWrapper.GetStiffenedPlates
        
        ' Make sure there is only one stiffened Plate
        Dim oPlatePart As IJPlatePart
        For Each oPlatePart In oStiffenedPlates
            Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
            Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
            Set oPlateWrapper.object = oPlatePart
            
            Dim oMfgPlatePart As IJMfgPlatePart
            ' Get the MfgPlate for the stiffened Plate
            If oPlateWrapper.PlateHasMfgPart(oMfgPlatePart) Then
                Dim eMoldedSide As enumPlateSide
                Dim eUpSide As enumPlateSide
                
                ' Get the plate MoldedSide and also Upside
                eMoldedSide = oMfgPlatePart.MoldedSide
                eUpSide = oMfgPlatePart.UpSide
                
                Dim oUVGeom3d As IJMfgGeom2d
                Dim oUVGeom2d As IJMfgGeom2d
                Dim oLocationGeom2d As IJMfgGeom2d
                Dim oMarkingInfoObj As Object
                
                ' Get the 3D UV mark
                Dim i As Long
                Dim oGeom2dCol As IJMfgGeomCol2d
                Dim oGeom3dCol As IJMfgGeomCol3d
               
                ' Get the 2d UV mark and also the plate location mark
                Set oGeom2dCol = oMfgProfilePart.FinalGeometriesAfterProcess2D
                
                If Not oGeom2dCol Is Nothing Then
                    If oGeom2dCol.Getcount > 0 Then
                        For i = 1 To oGeom2dCol.Getcount
                            Dim oGeom2d As IJMfgGeom2d
                           
                            Set oGeom2d = oGeom2dCol.GetGeometry(i)
                            
                            If ((oGeom2d.GetGeometryType = STRMFG_PLATELOCATION_MARK) Or (oGeom2d.GetGeometryType = STRMFG_UV_MARK)) Then
                                If (oGeom2d.GetGeometryType = STRMFG_UV_MARK) Then
                                    If oGeom2d.GetSubGeometryType = STRMFG_UV_MARK Then
                                        Set oUVGeom3d = oGeom2d
                                    Else
                                        Set oUVGeom2d = oGeom2d
                                    End If
                                Else
                                    Set oLocationGeom2d = oGeom2d
                                End If
                            End If
                            Set oGeom2d = Nothing
                        Next
                    End If
                End If
                
                Dim bstrEdgeXML As String
                Dim bstrPartDir As String
                Dim bThicknessUp As Boolean
                Set oMarkingInfoObj = oLocationGeom2d.SystemMark
                ' Get XML for the plate location mark
                oLocationGeom2d.OutputXML PROFILE_TYPE, bstrEdgeXML, oUVGeom2d, oUVGeom3d, , , oMarkingInfoObj
                
                Dim oGeomCrv As IJCurve
                Dim dStartX As Double, dEndX As Double, dDummy As Double
                
                Set oGeomCrv = oLocationGeom2d
                
                oGeomCrv.EndPoints dStartX, dDummy, dDummy, dEndX, dDummy, dDummy
                
                ' Get the Part Direction 'L' or 'R'
                bstrPartDir = Mid(bstrEdgeXML, (InStr(bstrEdgeXML, "PART_DIR=") + 10), 1)
                                
                ' If the location is mark is going from left to right (dStartX < dEndX)
                '       if the part direction is right, ThicknessUp is False.
                '       if the part direction is left, ThicknessUp is True.
                ' Else If the location is mark is going from right to left (dStartX > dEndX)
                '       if the part direction is right, ThicknessUp is True.
                '       if the part direction is left, ThicknessUp is False.
                ' Endif
                
                bThicknessUp = True
                If ((dStartX < dEndX) And (bstrPartDir = "R")) Then
                    bThicknessUp = False
                ElseIf ((dStartX > dEndX) And (bstrPartDir = "L")) Then
                    bThicknessUp = False
                End If
                
                ' If the Upside matches with MoldedSide and if the thickess direction is Up, then we should rotate the MfgPart by 180 degrees
                ' If the Upside does not match with MoldedSide and if the thickess direction is Down, then we should rotate the MfgPart by 180 degrees
                If ((eUpSide = eMoldedSide) And (bThicknessUp = True)) Then
                    GetPartRotationAngleForER = 180#
                ElseIf ((eUpSide <> eMoldedSide) And (bThicknessUp = False)) Then
                    GetPartRotationAngleForER = 180#
                End If
                
                Set oUVGeom3d = Nothing
                Set oUVGeom2d = Nothing
                Set oLocationGeom2d = Nothing
                Set oMarkingInfoObj = Nothing
                Set oGeom2dCol = Nothing
                Set oGeom3dCol = Nothing
                Set oGeomCrv = Nothing
            End If
            Set oMfgPlatePart = Nothing
            Set oPlateWrapper = Nothing
        Next
    End If
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

