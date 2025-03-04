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
'     Folder:  S:\StructManufacturing\Data\MfgRules\MfgPlateProcess
'     Project: MfgPlateProcess
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
Const MODULE = "MfgPlateProcess: "

Implements IJDMfgUnfoldParameters

Private Const sSOURCEFILE As String = "UserDefinedValues.cls"
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors   ' To collect and propagate the errors.
Private m_oError As IJEditError     ' Defined here for convenience

' Don't change the below values
Private Const EDGE_FEATURES As Long = &H1
Private Const CORNER_FEATURES As Long = &H10
Private Const SLOT_FEATURES As Long = &H100
Private Const SKETCHED_FEATURES As Long = &H1000
Private Const MOLDEDFORM_OPENINGS As Long = &H10000
' Don't change the above values

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
    Const METHOD As String = "UserDefinedValues: IJDMfgUnfoldParameters_GetDoubleValue"
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
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0001

    Case "LongProcessLine"
        'The bending line through the center of the plate is used for the Processing Line.
        'If another bending line is at least LongProcessLine longer,
        'this bending line becomes the Processing Line.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.25
        
    Case "KnuckleAngle"
        'When the base surface for a plate is defined by several faces,
        'and sinus to the angle between the tangents of two connecting faces are bigger than KnuckleAngle,
        'a knuckle line is defined between the two faces.
        'Magnitude of u cross v is sin w where u and v are the respective normals at the two faces
        'of the knuckle and w is the angle between u and v unit vectors. So, sin w = 0.001
        IJDMfgUnfoldParameters_GetDoubleValue = 0.001

    Case "KnuckleDist"
        'Two knuckle lines are considered to one line if endpoints on one line is
        'within this distance from the other line.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.001
        
    Case "SliverTolerance"
        ' If the input developable surafce has slivers within the given tolerance, they will be replaced with
        ' tolerant edges.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0035
        
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
        
    Case "DistEdgePoint"
        'To find knuckle lines, the surface normal is calculated at face edges.
        'When a point is found (by iteration) on an edge closer than this value to
        'the wanted point, the surface normal is calculated.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0001

    Case "UVInterval"
        'Calculating the surface normal at a face edge,
        'the point on the edge is found by iteration.
        'This iteration is stopped when the interval in
        'U- or V- direction is smaller than this value.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.00001

    Case "PointDist"
        'To unfold the curves on the plate, the curves are converted into points.
        'The distance between these points is defined by this value.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.01

    Case "EpsRollBoundary"
        'Tolerance for finding RollBoundaryLines.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0001

    Case "EpsFaceting"
        'Tolerance for faceting BendingLines
        IJDMfgUnfoldParameters_GetDoubleValue = 0.001
        
    Case "ChordHeight"
        'To unfold the curves on the plate, the curves are converted into points.
        'Number of points for a BSplineCurve are controlled by this value.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.005
            
    Case "MinimumCurveLengthForNesting"
        'Minimum Curve Length in the final output before going to nesting
        IJDMfgUnfoldParameters_GetDoubleValue = 0.001
        
    Case "ChordHeightToleranceForArcConversion"
        'An arc is converted to a line, if the chord height is less than this value
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0001
        
    Case "MaximumArcRadiusForArcConversion"
        ' If the maximum arc radius is more than this value, Arc will be broken into
        ' lines within ChordHeightTolerance.
        IJDMfgUnfoldParameters_GetDoubleValue = 999#
        
    Case "MaxOverlapDistForMarkingAndOuterContour"
        ' If the maximum distance between a marking line and the outer contour
        ' is less than this value, that marking line will be deleted.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.001
        
    Case "MaximumPCExtensionAlongEdgeForCollarPlates"
        'In case of collar plate Physical connection might not have extended all the way.
        'So this tolerance is to take care of such Physical connections.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.005
            
    Case "MaxCurveLengthToRemoveInSlotFeatures"
        ' Applying bevel offset corrections in slot feature curves might introduce tiny linesegments
        ' in the contour shape. If this length value property is set ( more than 0.0 ), we will
        ' remove tiny segments smaller than this value and adjust the neighboring curves
        ' by translating accordingly.
        IJDMfgUnfoldParameters_GetDoubleValue = 0#
        
    Case "MinimumFeatureSegmentLength"
        ' After imposing feature cuts, any resulting contour curves less than this value, will be removed from the group.
        ' Adjacent contours will be extended and trimmed to close the gap.
        IJDMfgUnfoldParameters_GetDoubleValue = 0#
                
    Case "ChordHeightToleranceForStrokingBsplineCurve"
        ' Given a B-spline curve and approximation tolerance, we generate number of points on the curve
        ' such that the line created by two adjacent points approximates the B-spline curve locally
        ' within this tolerance (chord-height < tolerance)
        IJDMfgUnfoldParameters_GetDoubleValue = 0.0001
        
    Case "BracketEdgeLengthForScalingMargin"
        ' When scaling margin is applied on a bracket plate, this option will help retaining the edge dimesion.
        ' If the edge adjacent to bracket edge on which margin is applied, is less than this specified value,
        ' it's dimesions will not be altered after scaling.
        ' Specifying 0.0 here will apply regular scaling to the entire shape, i.e. dimensions will be different after scaling.
        IJDMfgUnfoldParameters_GetDoubleValue = 0.1
        
    Case "VolumeLimitForSNUBox"
        ' When SNU_Box/SNU_FourSided option is used for unfolding shell plates, we construct a 3D box surrounding the plate surface.
        ' The hull surface is bounded with four reference planes of the bounding box.
        ' If the new bounding surface is much bigger than the original surface, SNU_Box option will be ignored and SNU algorithm is used instead.
        ' The ratio of the new bounding box volume to the original surface bounding box volume is more than this specified value,
        ' then SNU_Box will be considered as invalid.
        IJDMfgUnfoldParameters_GetDoubleValue = 2#

    Case "PanelContourBevel"
        ' This bevel will be applied on the manufacturing plate contour adjust by the panel margin.
        ' Value should be between 1 and 179 degrees and applies with respect to the mfg plate upside.
        ' Specify 90 degrees to remove the bevel ( Square Cut Bevel )
        IJDMfgUnfoldParameters_GetDoubleValue = 0#     ' Keep the bevel as is

    Case "ContourBevelForMargin"
        ' This bevel will be applied on the manufacturing plate or panel contour adjust by the non-panel margin.
        ' Value should be between 1 and 179 degrees and applies with respect to the mfg plate upside.
        ' Specify 90 degrees to remove the bevel ( Square Cut Bevel )
        IJDMfgUnfoldParameters_GetDoubleValue = 0#     ' Keep the bevel as is
        
    Case "MergeOrExtendDistanceForOpenings"
        ' If a marking line crosses a Opening and if the gap is less than this specified value in meters, it will be closed a line segment.
        IJDMfgUnfoldParameters_GetDoubleValue = 0#
        
    Case "MergeOrExtendDistanceForFeatures"
        ' If a marking line crosses a Feature and if the gap is less than this specified value in meters, it will be closed a line segment.
        IJDMfgUnfoldParameters_GetDoubleValue = 0#
        
    Case "OscillationCheckTolerance"
        ' When a plate is unfolded with SNU unfolding algorithm, the outer contour curves might have oscillations. The adjacent vectors between vertices might have sharp changes.
        ' It doesn't necessarily mean there is a problem with the output, but rather indicates a possible problem.
        ' If the oscillation amount ( in degrees ) for a curve is more than this value, MfgPlate will be put in to-do list.
        IJDMfgUnfoldParameters_GetDoubleValue = 10#
        
    Case "OuterGirthLengthCheckTolerance"
        ' When a plate is unfolded with SNU unfolding algorithm, some outer contour edges' unfolded girth lengths may not match with model 3D curve contour girth lengths.
        ' If the girth length mis-matching amount ( in mm ) for a curve is more than this value, MfgPlate will be put in to-do list
                ' Specifying a value of zero will effectively disable this check. This will be the default setting.
        IJDMfgUnfoldParameters_GetDoubleValue = 0#
        
    Case "MarkingGirthCheckTolerance"
        ' When a plate is unfolded with SNU unfolding algorithm, some curves might have Inner girth lengths differences before and after unfolding.
        ' This is the case where the marking line could go outside the plate contour because of bad unfolding process.
        ' If the marking end points for a curve are outside contour curve shape by more than this value ( in mm ), MfgPlate will be put in to-do list
                ' Specifying a value of zero will effectively disable this check. This will be the default setting.
        IJDMfgUnfoldParameters_GetDoubleValue = 0#
        
    Case Else
        IJDMfgUnfoldParameters_GetDoubleValue = -1
        
    End Select

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

Private Function IJDMfgUnfoldParameters_GetLongValue(ByVal bstrParameterList As String) As Long
    Const METHOD As String = "IJDMfgUnfoldParameters_GetLongValue"
    On Error GoTo ErrorHandler
    
    Dim bstrParameterName As String
    Dim sParameters() As String
    Dim lIndex As Long
    
    sParameters = Split(bstrParameterList, ":")
    
    bstrParameterName = sParameters(LBound(sParameters))
    If UBound(sParameters) > LBound(sParameters) Then
        lIndex = Int(sParameters(UBound(sParameters)))
    End If
    
    Select Case bstrParameterName
    Case "NumPoints"
        'Number of points calculated on intersection curve between plane and surface.
        'Used for finding bending lines (curvature of plate)
        IJDMfgUnfoldParameters_GetLongValue = 20

    Case "NFlatTest"
        'Number of points tested in U- and V- direction when testing plate for being plane.
        'NFlatTest*NFlatTest points are tested.
        IJDMfgUnfoldParameters_GetLongValue = 20
        
    Case "CloseOptionForFillingMarginGaps"
        ' Applying margin to an edge leave gaps between edge
        ' geometry and the adjacent edge geometry.
        ' This option tells the system how to connect the gap
        ' 1 - Edge geometry follows the natural/polynomial extension
        ' 2 - Gaps will be connected a new line segment
        ' 3 - Edge geometry follows the tangential extension
        ' Invalid Option - same as 1
        IJDMfgUnfoldParameters_GetLongValue = 2
 
    Case "CalculateSlantDimensions"
        ' If user set this flag to true ( 1 ), slant dimesnions will be caluclated for shell plates.
        ' If this value is false ( 0 ), slant dimensions will NOT be calculated.
        ' By default, this value is set to false ( 0 )
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "TrimMarkingLinesToBeforeChamferSurface"
        ' If user set this flag to true ( 1 ), marking will be trimmed to the upside surface without considering chamnfers.
        ' So in general a part of the marking lines is removed if the upside is the same side as the chamfer or
        'you could say It only marks on that surface of the part that is normal to the view direction.
        ' If this value is false ( 0 ), marking lines will be extended to all way to the plate edge.
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "SplitConstantBevelForCurvedEdges"
        ' Constant bevels are not split by default. As a result, entire entire with the constant bevel will be offset one value.
        ' If user wants more accurate bevel corrections, he needs to set SplitConstantBevelForCurvedEdges parameter to 1, which tells
        ' the system to split edge even if it is a constant bevel.
        ' 1 - Split Constant Bevels for curved edges
        ' 0 - Do not Split Constant Bevels for curved edges
        IJDMfgUnfoldParameters_GetLongValue = 0
        
''''''    Case "MarksOverLappingPrecedence"
''''''        Select Case lIndex
''''''            Case 1:
''''''                IJDMfgUnfoldParameters_GetLongValue = STRMFG_PLATELOCATION_MARK
''''''            Case 2:
''''''                IJDMfgUnfoldParameters_GetLongValue = STRMFG_PROFILELOCATION_MARK
''''''            Case 3:
''''''                IJDMfgUnfoldParameters_GetLongValue = STRMFG_FRAMELINE_MARK
''''''            Case 4:
''''''                IJDMfgUnfoldParameters_GetLongValue = STRMFG_WATERLINE_MARK
''''''            Case 5:
''''''                IJDMfgUnfoldParameters_GetLongValue = STRMFG_BUTTOCKLINE_MARK
''''''            Case 6:
''''''                IJDMfgUnfoldParameters_GetLongValue = STRMFG_TemplateLocationMarkLine
''''''            Case Else:
''''''                IJDMfgUnfoldParameters_GetLongValue = -1
''''''        End Select
    
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
                
    Case "MarginsShrinkagesOrder"
        ' Margin and shrinkages are applied on manufacturing part in post-unfolding stage.
        ' With this parameter, user can control, in which order margin and shrinkages are applied
        ' 0 - Random Order
        ' 1 - Apply Margins, ShrinkageMargins and Shrinkages in that order
        ' 2 - Apply ShrinkageMargins, Margins and Shrinkages in that order
        ' 3 - Apply Shrinkages, Margins and ShrinkageMargins in that order
        ' 4 - Apply Shrinkages, ShrinkageMargins and Margins in that order
        IJDMfgUnfoldParameters_GetLongValue = 1
                
    Case "ApproximateFeatureSplineWithLineOrArc"
        ' With this parameter, user has ability to change the bspline curves in a feature objects to a line or arc.
        ' 0 - Keeep the bspline curves as is
        ' 1 - Convert bspline curves to arc, if possible.
        ' 2 - Convert bspline curves to line, if possible.
        ' 3 - Convert bspline curves to line or arc, if possible.
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "ShiftFeaturesWhenNoScalingForMargin"
       ' When a margin is applied "As Shrinkage" type, and if a openining/feature has no scaling applied, this option comes into play.
       ' Opening/Margin can be shifted with the margin value, Or the distance can be scaled w.r.to the edge.
       ' This property tells which objects need to be shifted.
       ' Example IJDMfgUnfoldParameters_GetLongValue = (EDGE_FEATURES Or CORNER_FEATURES Or SLOT_FEATURES)
       IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "ApplyScalingOnOpenings"
       'With this parameter, user has ability to control how scaling is applied on MoldedForm Openings.
       'If the flag is set to 0, we do not apply scaling to Openings at all.
       'If the flag is set to 1, we always apply scaling to Openings.
       'If the flag is set to 2, we apply scaling to Openings for normal shrinkage, but No scaling of Openings in scaling margin case.
       'If the flag is set to 3, feature scaling is done for scaling margin, but not for regular shrinkage.
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "ApplyScalingOnFeatures"
       'With this parameter, user has ability to control how scaling is applied on features ( Sketched features, slots, scallops, corner features, etc.. ).
       'If the flag is set to 0, we do not apply scaling to features at all ( current behavior ).
       'If the flag is set to 1, we always apply scaling to features.
       'If the flag is set to 2, we apply scaling to features for normal shrinkage, but No scaling of features in scaling margin case.
       'If the flag is set to 3, feature scaling is done for scaling margin, but not for regular shrinkage.
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "MergeChildFeatureContours"
        ' If a struct feature is child of another parent feature, in the final XML,
        ' both the parent edge and child edges will be merged as a single edge, based on this flag.
        ' 0 - Do not merge. Each feature edge will remain separate SMS_EDGE.
        ' 1 - Merge parent and child feature edges into one single SMS_EDGE.
        IJDMfgUnfoldParameters_GetLongValue = 0

    Case "CreateBsplineCurvesInUnfold"
        ' Our 3rd party SNU unfolding algorithm outputs the unfolded curves as line strings.
        ' With this new parameter rule, user has the option of converting the unfolded output to a bspline curve, rather than a line string.
        ' 0 - No Change in behavior. Output will remain as a line string.
        ' 1 - Unfolded Output will be converted a bspline curve.
        IJDMfgUnfoldParameters_GetLongValue = 0

    Case "LineArcApproximationAlgorithm"
        ' While getting XML data for part monitor command, each curve in the manufacturing output goes
        ' through a line/arc approximation algorithm to convert the curves to lines and arcs.
        ' With this process parameter, user can control which line/arc approximate to use.
        ' 1 - GTYPE Line/Arc Approximation Algorithm implemented by using GTYPE curves
        ' 2 - ACIS Line Arc Approximation Algorithm implemented by using ACIS Edges
        IJDMfgUnfoldParameters_GetLongValue = 1

    Case "BreakGeometriesAtCenterLineForVaryingBevel"
        ' If a physical connection is crossing the center line for transversal bulkhead plates,
        ' varying bevel is not generating symmetric output. With this parameter, user has ability to
        ' split the edge geometry at the center line location so that port and starboard curves
        ' are handled separately to give symmetric output.
        ' 0 - Do Not Split at the center line ( Y=0 ) location
        ' 1 - Split at the center line ( Y=0 ) location
        IJDMfgUnfoldParameters_GetLongValue = 1
        
    Case "AccumulateShrinkages"
        ' If there are multiple shrinkages on a plate part, this parameter controls how to accumulate them.
        ' This is a number which is pointing to an SRD rule
        ' 151 - Add shrinkages
        ' 152 - Multiply Shrinkages
        ' 153 - Average of Shrinkages
        IJDMfgUnfoldParameters_GetLongValue = 151
        
        Dim oMfgPlatePartWrapper As New MfgRuleHelpers.PlatePartHlpr
        Dim oMfgPlate As IJMfgPlatePart
        
        Set oMfgPlatePartWrapper.object = m_oObject
        
        If oMfgPlatePartWrapper.PlateHasMfgPart(oMfgPlate) Then
            If oMfgPlate.PanelMode = True Then
                ' 153 - Average of Shrinkages
                IJDMfgUnfoldParameters_GetLongValue = 153
            End If
        End If
        
    Case "FeatureMoveDirection"
        ' This flag controls how to move the feature curves while applying margin.
        ' 0 - Move feature geometry normal to the margin edge.
        ' 1 - Move feature geometry along the bounding non-feature edge
        IJDMfgUnfoldParameters_GetLongValue = 1
            
    Case "PartOutputLocationOption"
        ' With this option, user will have ability to move the manufacturing part output to desired location
        ' 0 - No change in output. Part XML reference position will be decided by the software. This is default output.
        ' 1 - Align Centroid of the outer contour with origin ( 0, 0 )
        ' 2 - Align mid point of the Outer contour's range box with origin
        ' 3 - Align lowest left point of the Outer contour with origin
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "PartAlignmentOption"
        ' With this option, user will have ability to rotate the part data in XML file and in part monitor command.
        ' This option will have impact only while exporting the data to part monitor or to nesting
        ' 0 - No change in output. Part rotation will be decided by the software. This is default output.
        ' 1 - Align longest edge to the closest X-axis or Y-Axis
        ' 2 - Align longest edge to the X-axis
        ' 3 - Align Minimum Bounding Box's longest edge to the closest X-axis or Y-Axis
        ' 4 - Align Minimum Bounding Box's longest edge to the X-axis
        IJDMfgUnfoldParameters_GetLongValue = 0
                
    Case "BuiltUpPartAlignmentOption"
        ' With this option, user will have ability to rotate the BuiltUp part data in XML file and in part monitor command.
        ' This option will have impact only while exporting the 'Web Plate' type data to part monitor or to nesting
        ' 0 - PartAlignmentOption will have the priority.
        ' 1 - Align the web plate so that landing curve will be on bottom side.
        ' 2 - Align the web plate so that landing curve will be on top side.
        
        IJDMfgUnfoldParameters_GetLongValue = 0
        
        If TypeOf m_oObject Is IJPlatePart Then
            Dim oMfgPlateWrapper As MfgRuleHelpers.PlatePartHlpr
            Dim oMfgPlatePart As IJMfgPlatePart
            
            Set oMfgPlateWrapper = CreateObject("MfgRuleHelpers.PlatePartHlpr")
            Set oMfgPlateWrapper.object = m_oObject
            
            If oMfgPlateWrapper.PlateHasMfgPart(oMfgPlatePart) Then
                Dim sNamingCategory As String
                oMfgPlatePart.GetPlateNamingCategory sNamingCategory
                If sNamingCategory = "WB" Then
                    ' IJDMfgUnfoldParameters_GetLongValue = 1
                End If
            End If
        End If
                
    Case "PanelAlignmentOption"
        ' With this option, user will have ability to rotate the Panel part data in XML file and in part monitor command.
        ' This option will have impact only while exporting the Panel data to part monitor or to nesting
        ' 0 - No change in output. Part rotation will be decided by the software. This is default output.
        ' 1 - Align with most stiffener markings with X axis ( flow direction )
        ' 2 - Align with most stiffener markings with Y axis ( flow direction )
        ' 3 - Align with most inner seam markings with X axis
        ' 4 - Align with most inner seam markings with Y axis
        ' 5 - Align so that Assembly Primary Orientation with Unfolded X axis ( +/- not considered )
        ' 6 - Align so that Assembly Primary Orientation with Unfolded Y axis ( +/- not considered )
        ' 7 - Align so that Assembly Secondary Orientation with Unfolded X axis ( +/- not considered )
        ' 8 - Align so that Assembly Secondary Orientation with Unfolded Y axis ( +/- not considered )
        ' 9 - Align so that Assembly Orientation matches with Unfolded XY-axes as much as possible
        IJDMfgUnfoldParameters_GetLongValue = 0

    Case "ReverseBevelPropsIfNormalAreOpposite"
        ' When Base and Offset ports are flipped between two adjacent plate parts, manufacturing could detect that the normals were different,
        ' then reverse non-reference part bevels.
        ' 0 - Do not flip the non-reference part bevel properties
        ' 1 - Flip the non-reference part bevel properties
        IJDMfgUnfoldParameters_GetLongValue = 1
            
        Case "SpecialExtendAndTrimForCollarPlates"
        ' As part of CR#129420, Some customers want special processing of collar plates while applying bevel corrections.
        ' They want smooth continuous curves, even though the collar edges may not have bevels.
        ' 0 - No special processing for collar plate
        ' 1 - Apply special extend and trim functionality for collars.
        IJDMfgUnfoldParameters_GetLongValue = 1
                
    Case "PartialRecompute"
        ' For planar plates, we have implemented partial recompute functionality long time back.
        ' This flag provides the ability to turn on/off the Partial Recompute Functionality.
        ' 0 - Turn Off Partial Recompute Functionality
        ' 1 - Turn On Partial Recompute Functionality
        IJDMfgUnfoldParameters_GetLongValue = 1
                
    Case "ConvertCurveToPrimitivesBeforeUnfolding"
        ' For developable plates ( like knuckle plates, revolved plates, flanged brackets, corrugate plates and linear extruded plates ),
        ' we convert the input boundary curves and marking to primitives ( lines and arcs ) before unfolding.
        ' This flag provides the ability to turn on/off that conversion process.
        ' 0 - Do NOT convert the boundary curves and marking to primitive lines and arcs
        ' 1 - Convert the boundary curves and marking to primitive lines and arcs
        ' 2 - if the plate is a FlangedBracket, then do NOT convert. This will make sure that original arc radii is maintained after unfolding.
        '     For all the other plate types, convert to primitives.
        IJDMfgUnfoldParameters_GetLongValue = 1
                
    Case "PanelEndMarginCodeListNumber"
        ' This codelist number specifies the margin type codelist value assigned for panel end margin.
        ' System uses this value to separate the panel end margin from the other margin types
        ' and apply the panel margin after all other margins/shrinkages are applied.
        IJDMfgUnfoldParameters_GetLongValue = 9
                
    Case "MergeOrExtendOptionForOpenings"
        ' If a marking line touches an Opening, this option control the extend/trim behaviour
        '    0 - Apply MergeOrExtendDistanceForOpenings for all cases
        '    1 - Apply MergeOrExtendDistanceForOpenings only for inner holes
        '    2 - Apply MergeOrExtendDistanceForOpenings only for openings impacting the outer contour
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "MergeOrExtendOptionForFeatures"
        ' If a marking line touches a Feature object, this option control the extend/trim behaviour
        '    0 - Apply MergeOrExtendDistanceForOpenings for all cases
        '    1 - Apply MergeOrExtendDistanceForOpenings only for hole features
        '    2 - Apply MergeOrExtendDistanceForOpenings only for features impacting the outer contour
        IJDMfgUnfoldParameters_GetLongValue = 0
                        
    Case "ApplyGeometryBevelOnEdges"
        ' If a contour/feature doesn't have any PC/FET associated with it, this parameter will let the system compute long-point bevel automatically.
        '    0 - Do Not apply any Geometry based varying bevel on edge
        '    1 - Apply Geometry based varying bevel only on Feature Edges
        '    2 - Apply Geometry based varying bevel only on Contour Edges
        '    3 - Apply Geometry based varying bevel on Contour and Feature Edges
        IJDMfgUnfoldParameters_GetLongValue = 0
        
    Case "ComputePanelCassettePoints"
        ' For a panel, reference points on profile part are needed for cassetting placement.
        ' With this option, user can control if they want system to compute and get the cassette points in XML.
        '    0 - Do Not compute Cassette Points
        '    1 - Compute Cassette Points
        IJDMfgUnfoldParameters_GetLongValue = 0
                
    Case "DeleteZeroParameterBevels"
        ' If there is any physical connection having invalid parameters, this parameter controls whether to delete that manufacturing bevel or not.
        '    0 - Do Not delete bevels having all zero depth parameters
        '    1 - Delete bevels having all zero depth parameters
        IJDMfgUnfoldParameters_GetLongValue = 0
                
    Case "FormingOptionForSNUUnfold"
        ' There are several ways of bending shell plates. This option specifies how to unfold a shell plate considering the manufacturing forming method.
        ' With Cold forming, Plate is bent at normal temperatures in a bending machine. Material stretches after bending so Unfolded Plate Dimensions will be smaller.
        ' With Themal Forming, Heating of plate before being bent. Material shrinks after cooling so Unfolded Plate Dimensions will be larger
        '    0 - Thermal Bending
        '    1 - Cold Bending
        '    2 - No Deformation
        IJDMfgUnfoldParameters_GetLongValue = 0

    Case "StoreBestFitDimensions"
        ' For a manufacturing plate, the persisted Length and Width properties on IJMfgPlatePart interface are computed by getting the global x-range and y-range differences.
        ' However, these may not be optimal dimensions for the unfolded plate.
        ' With this option, User has an ability to get the best-fit length and width values, which do not change depending on the MfgPlate orientation
        '    0 - Compute length and width valued based on Global ranges
        '    1 - Compute best-fit length and width valued based on minimum rectangle surrounding mfg part shape
        IJDMfgUnfoldParameters_GetLongValue = 0

    Case Else
        IJDMfgUnfoldParameters_GetLongValue = -1
        
    End Select

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1055, , "RULES")
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


