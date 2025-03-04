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
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module:UserDefinedValues
'
' Description:  Determines the process settings for the mfg profile
'
' Author:
'
' Comments:
' 8 Feb, 2008    Creation
'*******************************************************************************
Option Explicit
Const MODULE = "MfgMemberProcess"

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
        IJDMfgUnfoldParameters_GetDoubleValue = 0.005

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

    Case Else
        IJDMfgUnfoldParameters_GetDoubleValue = -1

    End Select

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2034, , "RULES")
End Function

Private Function IsTubularCrossSection(oMemberPart As ISPSMemberPartCommon) As Boolean
    Dim XnTypeName As String
    XnTypeName = oMemberPart.CrossSection.CrossSectionTypeName

    If XnTypeName = "RoundBar" Or XnTypeName = "HSSC" Or _
       XnTypeName = "RoundTube" Or XnTypeName = "PIPE" Then
        IsTubularCrossSection = True
    Else
        IsTubularCrossSection = False
    End If
End Function

Private Function IJDMfgUnfoldParameters_GetLongValue(ByVal bstrParameterName As String) As Long
    Const METHOD As String = "IJDMfgUnfoldParameters_GetLongValue"
    On Error GoTo ErrorHandler

    Select Case bstrParameterName
    Case "NumPoints"
        'Number of points calculated on intersection curve between plane and surface.
        'Used for finding bending lines (curvature of plate)
        IJDMfgUnfoldParameters_GetLongValue = 20

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

    Case "MarginByOffset"
        ' There are two different ways to apply margin. One is by Translation and another by applying constant
        ' offset through out the length of the curve. By default, for all profiles/members margin is applied
        ' by translation ( i.e. contour curve will be shifted in the margin direction ). If user wants to apply
        ' margin by Translation, they need to change the following value to 0.
        ' 0 - Translation 1 - Offset
        IJDMfgUnfoldParameters_GetLongValue = 1

    Case "ApplyScalingOnFeatures"
       'With this parameter, user has ability to control how scaling is applied on features ( Sketched features, slots, scallops, corner features, etc.. ).
       'If the flag is set to 0, we do not apply scaling to features at all ( current behavior ).
       'If the flag is set to 1, we always apply scaling to features.
       'If the flag is set to 2, we apply scaling to features for normal shrinkage, but No scaling of features in scaling margin case.
       'If the flag is set to 3, feature scaling is done for scaling margin, but not for regular shrinkage.
        IJDMfgUnfoldParameters_GetLongValue = 0

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

    Case "ApplyGeometryBevelOnEdges"
        ' If a feature doesn't have any PC/FET associated with it, this parameter will let the system compute long-point bevel automatically.
        '    0 - Do Not apply any Geometry based varying bevel on edge
        '    1 - Apply Geometry based varying bevel on Feature Edges
        IJDMfgUnfoldParameters_GetLongValue = 0

    Case "MacroLibraryIdentifier"
        ' This flag specifies the identifier for the library of macro definitions.
        ' This identifier is used to locate the record in the JMfgMacroDefOption catalog view.
        ' OOTB rules follow the convention that Tubular members have this ID as 101 and
        ' Members that do not have circular cross-section have this ID as 102.
        If IsTubularCrossSection(m_oObject) Then
            IJDMfgUnfoldParameters_GetLongValue = 101
        Else
            IJDMfgUnfoldParameters_GetLongValue = 102
        End If

    Case "TubularMemberUpside"
        ' This flag allows control of the "View" in Part monitor for a tubular member
        '   0 - System default ('zero' parameter of cross-section)
        '   1 - Member seam (NB: Ensure it is present!)
        '   2 - Opposite member seam (NB: Ensure it is present!)
        '   3 - Any inner tube material extremity
        '   4 - Any outer tube material extremity

        Dim oMemberPart As ISPSMemberPartPrismatic
        Set oMemberPart = m_oObject

        Dim MemberSeamInfo As ISPSMemberPrismaticSeamData
        Set MemberSeamInfo = oMemberPart.SeamData

        If Not MemberSeamInfo Is Nothing Then
            IJDMfgUnfoldParameters_GetLongValue = 2
        Else
            IJDMfgUnfoldParameters_GetLongValue = 3
        End If

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

