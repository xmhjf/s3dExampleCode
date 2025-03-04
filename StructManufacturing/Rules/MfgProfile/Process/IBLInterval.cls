VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBLInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module:
'
' Description:  Determines the process settings for the mfg profile
'
' Author:
'
' Comments:
'   The measurement points are every 500mm apart with the length from the left end of the profile. However it will be controlled by the profile length
'   Profile Length l >1,000mm:  Every 500mm
'   Profile Length 500=< l<999mm:   Every 250mm
'   Profile Length 250=< l<499mm:?      Every 125mm
'   Profile Length l < 249mmNo IBL
'   Request to customize the divided interval individually at each profile to considerate a curvature of profiles from 500, 250 and 125.
'   IBL must maintain minimum of 15mm from the above landing curve edge
'   IBL also must maintain minimum of 40mm from profile flange or top
'   Maximum number of IBL's is 5
'   If multiple IBLs are needed, curves will overlap by 3 divided points. (In case of 500mm, 1,000mm)
'   Intersection points of IBL's must be the divisible location
'   If multiple IBLs are placed center of over lapped divided line should be marked. Also length of 3D shape measure bottom end to end in below red arrow line in drawing and also height of marked center of divided line normal to red arrow line in green line of below figure

'*******************************************************************************
Option Explicit

Implements IJDMfgProfileBendingLineRule
Const MODULE = "MfgProfileProcess"
Private Sub IJDMfgProfileBendingLineRule_GetBendingLineParameters(ByVal oProfilePart As Object, pBottomMargin As Double, pTopMargin As Double, pStartMargin As Double, pEndMargin As Double, pOverlapDist As Double, pMaxNoOfBendingLines As Long)
    Const METHOD = "IBLSingle: IJDMfgProfileBendingLineRule_GetBendingLineParameters"
    On Error GoTo ErrorHandler
    
    Dim oMfgProfilePart As IJMfgProfilePart
    
    Dim oProfilePartHlpr As MfgRuleHelpers.ProfilePartHlpr
    Set oProfilePartHlpr = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfilePartHlpr.object = oProfilePart
                        
    If oProfilePartHlpr.ProfileHasMfgPart(oMfgProfilePart) = False Then
        Exit Sub
    End If
    
    Dim oLandCrvWireBody As IJDModelBody
     Dim dLandCrvLength As Double, dDummy As Double
    
    Set oLandCrvWireBody = oMfgProfilePart.LandingCurve
    oLandCrvWireBody.GetDimMetrics 0.000001, dDummy, dLandCrvLength, dDummy, dDummy
    
    If dLandCrvLength < 0.25 Then
        oMfgProfilePart.IntervalDistance = 0#
    ElseIf dLandCrvLength < 0.5 Then
        oMfgProfilePart.IntervalDistance = 0.125
    ElseIf dLandCrvLength < 1# Then
        oMfgProfilePart.IntervalDistance = 0.25
    Else
        oMfgProfilePart.IntervalDistance = 0.5
    End If
    
    Dim oProfileSection As IJDProfileSection
    Set oProfileSection = oProfilePart
    
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = oProfileSection.CrossSection
    
    Dim oCSPartClass As IJDCrossSectionPartClass
    Set oCSPartClass = oCrossSection.GetPartClass
    
    Dim SectionName As String
    SectionName = oCSPartClass.CrossSectionTypeName
    
    Dim IBLmarginNotSet As Boolean
    IBLmarginNotSet = True
    
    If UCase(Left(SectionName, 4)) = "BULB" Then
        Dim oAttr As IJDAttribute
        Set oAttr = GetAttribute(oCrossSection, "IJUAXSectionFlange", "FilletRadius")
        
        If Not oAttr Is Nothing Then
            oMfgProfilePart.BottomMargin = 0.005
            oMfgProfilePart.TopMargin = 5 * CDbl(oAttr.Value)
            
            oMfgProfilePart.FlangeBottomMargin = 0.001
            oMfgProfilePart.FlangeTopMargin = 0.001
            
            IBLmarginNotSet = False
        End If
    End If
    
    If IBLmarginNotSet Then
        oMfgProfilePart.BottomMargin = 0.015
        oMfgProfilePart.TopMargin = 0.04
    End If
    
    oMfgProfilePart.StartMargin = 0
    oMfgProfilePart.EndMargin = 0
    
    oMfgProfilePart.OverlapSide = eOverLapSideNeutral
    oMfgProfilePart.OverlapDistance = 2# * (oMfgProfilePart.IntervalDistance)
    
    oMfgProfilePart.MaxNoOfBendingLines = 5
    
    oMfgProfilePart.NextIBLStartSide = eNextBendLineStartAtSame
    oMfgProfilePart.IntersectPtsOnInterval = True
        
    oMfgProfilePart.SetBendingLinePos StartAtBottom, 0#
    ' oMfgProfilePart.LastIBLoption = StartAtBottom ' CR 343203

    ' Inverse Bending Line Control Marks
    '   0 -- Do not generate Control marks
    '   1 -- Generate control marks midway between
    '        end of one IBL and start of other IBL
    '   2 -- Generate Control marks where IBLs intersect
    oMfgProfilePart.GenerateIBLControlMarks = 2 ' Generate Control marks where IBLs intersect
    
    oMfgProfilePart.GenerateOneIBLIfPossible = True
    ' oMfgProfilePart.UserPriorityForSingleIBL = "LT-RT,LN-RN,LB-RB,LT-RN,LT-RB,LN-RT" ' CR 292217

    ' Use the profile part as input
    Dim oProfileClass As Object
    If TypeOf oProfilePart Is IJStiffenerPart Then
        Set oProfileClass = New StructDetailObjects.ProfilePart
    ElseIf TypeOf oProfilePart Is IJBeamPart Then
        Set oProfileClass = New StructDetailObjects.BeamPart
    End If
    Set oProfileClass.object = oProfilePart
    
    ' Check if the cross-section is a built-up
    Dim bIsBuiltUp As Boolean
    bIsBuiltUp = oProfileClass.IsCrossSectionABuiltUp
    
    If bIsBuiltUp = False Then
        ' If it's not a built-up, then we adjust the TopMargin to avoid painting where the flange is
        oMfgProfilePart.TopMargin = 0.04 + oProfileClass.FlangeThickness
    End If
        
    'adjust the active region for margins
    Dim oStructConnectable As IJStructConnectable
    
    If TypeOf oProfilePart Is IJStructConnectable Then
        Dim oBasePort As IJPort
        Dim oOffsetPort As IJPort
        Dim oLateralPortsCol As IJElements
        
        Set oStructConnectable = oProfilePart
        oStructConnectable.GetBaseOffsetLateralPorts vbNullString, False, oBasePort, oOffsetPort, oLateralPortsCol

        If Not oBasePort Is Nothing Then
            oMfgProfilePart.StartMargin = -1 * GetMarginValueOnThePort(oBasePort)
        End If
        If Not oOffsetPort Is Nothing Then
            oMfgProfilePart.EndMargin = -1 * GetMarginValueOnThePort(oOffsetPort)
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

