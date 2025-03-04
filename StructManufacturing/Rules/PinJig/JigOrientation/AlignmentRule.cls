VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlignmentRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgPinJigOrientation
' Module: AlignmentRule
'
' Author: Manasa Jaisetty
'
' Comments: This rule drives the Pinbed orientation. It proposes the PinJig alignment
'*******************************************************************************
Option Explicit
Private Const MODULE = "StrMfgPinJigOrientation.AlignmentRule"

Implements IJDPinJigAlignmentRule
Private Enum MfgJigAlignmentType
    AlignToDefault = 1
    AlignToLongestEdge = 2
    AlignToGlobal = 3
    AlignToPinsOnCenterLine = 4
    AlignToLeftContour = 5
    AlignToBottomContour = 6
    AlignToRightContour = 7
    AlignToTopContour = 8
    'AlignToNewType5 = 9'For further customization
    'AlignToNewType5 = 10
    'AlignToCustomType1 = 10001
    'AlignToCustomType2 = 10002
    'AlignToCustomType3 = 10003
    'AlignToCustomType4 = 10004
    'AlignToCustomType5 = 10005
End Enum

Private Function IJDPinJigAlignmentRule_GetAllowableAlignmentTypes(ByVal pSupportedPlatesColl As IJElements) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
    Const METHOD = "IJDPinJigAlignmentRule_GetAllowableAlignmentTypes"
    On Error GoTo ErrorHandler
    
    'If any of supported the plates has a centerline passing through it,
    'then add "Based on Pins On Axis" else return all the other types
    Dim oAlignmentTypes As Collection
    Set oAlignmentTypes = New Collection
    
    oAlignmentTypes.Add AlignToDefault
    oAlignmentTypes.Add AlignToLongestEdge
    oAlignmentTypes.Add AlignToGlobal
    
    oAlignmentTypes.Add AlignToLeftContour
    oAlignmentTypes.Add AlignToBottomContour
    oAlignmentTypes.Add AlignToRightContour
    oAlignmentTypes.Add AlignToTopContour
        
    If ArePlatesHavingCenterLineIntersection(pSupportedPlatesColl) = True Then
        oAlignmentTypes.Add AlignToPinsOnCenterLine
    End If
    
    Set IJDPinJigAlignmentRule_GetAllowableAlignmentTypes = oAlignmentTypes
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5026, , "RULES")
End Function

Private Function IJDPinJigAlignmentRule_GetDefaultAlignmentType(ByVal pSupportedPlatesColl As IJElements) As Long
    Const METHOD = "IJDPinJigAlignmentRule_GetDefaultAlignmentType"
    On Error GoTo ErrorHandler
      
     'UNCOMMENT THE BELOW LINES TO RETURN "PinsOnCenterLine" FOR CENTERLINE PLATES
     'If ArePlatesHavingCenterLineIntersection(pSupportedPlatesColl) = False Then
        'Return the default alignment type
        IJDPinJigAlignmentRule_GetDefaultAlignmentType = AlignToDefault
     'Else
        'IJDPinJigAlignmentRule_GetDefaultAlignmentType = AlignToPinsOnCenterLine
     'End If
    

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5026, , "RULES")
End Function

'Description: This method would return the adjusted pin bed positions in the following order: OriginPos, XPos, OtherPos, YPos
'             It would adjust the pin bed corners based on the different tolerances(such as Offset,MinOverHang,additional row/col) provided by the rule.
'             The corners would also be ordered to get the origin position at the desired corner
Private Function IJDPinJigAlignmentRule_GetAdjustedPinBedPositions(ByVal pPinJig As Object) As IMSCoreCollections.IJElements
    Const METHOD = "IJDPinJigAlignmentRule_GetAdjustedPinBedPositions"
    On Error GoTo ErrorHandler
    
    'Get PinJig base plane
    Dim oPinJig As IJPinJig
    Set oPinJig = pPinJig
    Dim RootX As Double, RootY As Double, RootZ As Double
    Dim NormX As Double, NormY As Double, NormZ As Double
    oPinJig.GetBasePlane NormX, NormY, NormZ, RootX, RootY, RootZ
    
    Dim oBasePlane As IJPlane
    Set oBasePlane = New Plane3d
    oBasePlane.DefineByPointNormal RootX, RootY, RootZ, NormX, NormY, NormZ
    
    'Get PinJig supported plates
    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates
    
    'Get PinJig remarking surface
    Dim oRemarkingSurf As IJSurfaceBody
    Set oRemarkingSurf = oPinJig.RemarkingSurface
    
    'Get the transversal and longitudinal interval
    Dim dTransIntvl As Double, dLongiIntvl As Double
    oPinJig.GetSpacing dTransIntvl, dLongiIntvl
    
    'Get the Offset/Overhang values
    Dim dTransMinOverHang As Double, dLongiMinOverHang As Double
    Dim dTransOffset As Double, dLongiOffset As Double
    Dim dAddRow As Double, dAddCol As Double
    
    dTransMinOverHang = IJDPinJigAlignmentRule_SpecifyValueForPurpose(TransverseMinimumOverhang, pPinJig, dTransIntvl)
    dLongiMinOverHang = IJDPinJigAlignmentRule_SpecifyValueForPurpose(LongitudinalMinimumOverhang, pPinJig, dLongiIntvl)
    dTransOffset = IJDPinJigAlignmentRule_SpecifyValueForPurpose(TransverseOffsetForOriginPin, pPinJig, dTransIntvl)
    dLongiOffset = IJDPinJigAlignmentRule_SpecifyValueForPurpose(LongitudinalOffsetForOriginPin, pPinJig, dLongiIntvl)
    dAddRow = IJDPinJigAlignmentRule_SpecifyValueForPurpose(AddAdditionalRow, pPinJig, dTransIntvl)
    dAddCol = IJDPinJigAlignmentRule_SpecifyValueForPurpose(AddAdditionalColumn, pPinJig, dLongiIntvl)
        
    Dim oContourLines As IJElements
    Dim oJigFloorUtil As New PinJigFloorUtils
    Dim oCorners As IJElements
    Dim oOrderedCorners As IJElements
    
    ''Note: STRMFG_PinJig_Adjustment_Extension will have no meaning when STRMFG_PinJig_Adjustment_BoxPoints is based on CONTOURS.
    
    '' ////////////Allowed Possible Values //////////////////////////////////////////////////////////
    'eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_Contours Or STRMFG_PinJig_Adjustment_Offset_Average
    'eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_Contours Or STRMFG_PinJig_Adjustment_Offset_Fixed
    '
    'eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_Corners Or STRMFG_PinJig_Adjustment_Offset_Average Or STRMFG_PinJig_Adjustment_Extension_Yes
    'eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_Corners Or STRMFG_PinJig_Adjustment_Offset_Average Or STRMFG_PinJig_Adjustment_Extension_No
    'eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_Corners Or STRMFG_PinJig_Adjustment_Offset_Fixed Or STRMFG_PinJig_Adjustment_Extension_Yes
    'eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_Corners Or STRMFG_PinJig_Adjustment_Offset_Fixed Or STRMFG_PinJig_Adjustment_Extension_No
    '////////////////////////////////////////////////////////////////////////////////////////////////
    Dim eAdjustmentType As StrMfgPinJigAdjustmentTypes
    Dim eOriginPos As StrMfgPinJigOriginPosition
    Dim eDistanceMeasureType As MfgJigAlignmentDistanceMeasureType
    
    'Get the pin alignment type.
    Dim lPinJigAlignmentType As Long
    lPinJigAlignmentType = oPinJig.PinJigAlignmentType
    
    Select Case lPinJigAlignmentType
    
        Case AlignToDefault
            
            'NOTE: STRMFG_PinJig_Adjustment_BoxPoints_Corners is NOT APPLICABLE for this alignment option
            eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_Contours Or STRMFG_PinJig_Adjustment_Offset_Fixed
            eOriginPos = oPinJig.PinOriginPosition
            
            Set oContourLines = GetPinJigContourLines(oPinJig, ProjectedJigRemarking)
            Set oCorners = oJigFloorUtil.GetMinimumRectangleOfProjectedContours(oContourLines, oBasePlane)
                        
            Set oOrderedCorners = oJigFloorUtil.OrderPointsBasedOnAFLU(oContourLines, oCorners, oBasePlane, eOriginPos)
            
            Set IJDPinJigAlignmentRule_GetAdjustedPinBedPositions = oJigFloorUtil.GetAdjustedPinBedCorners(oOrderedCorners, oContourLines, oBasePlane, eAdjustmentType, dTransOffset, dLongiOffset, dTransMinOverHang, dLongiMinOverHang, dAddRow, dAddCol, dTransIntvl, dLongiIntvl)
        
        Case AlignToLongestEdge
            
            eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_Corners Or STRMFG_PinJig_Adjustment_Offset_Fixed Or STRMFG_PinJig_Adjustment_Extension_Yes
            eOriginPos = oPinJig.PinOriginPosition
            eDistanceMeasureType = ProjectedJigStraightLength
            
            ' Below are the various options for determining the longest edge. Get the corresponding contour lines in the above step
            ' MfgJigAlignmentDistanceMeasureType --
            '   ProjectedJigStraightLength = 0
            '   ProjectedJigGirthLength = 1
            '   JigSurfaceStraightLength = 2
            '   JigSurfaceGirthLength = 3
            Dim oSurfaceContourLines As IJElements
            Set oSurfaceContourLines = GetPinJigContourLines(oPinJig, JigSurfaceRemarking)
            Set oContourLines = GetPinJigContourLines(oPinJig, ProjectedJigRemarking)
            
            Set oCorners = oJigFloorUtil.GetBoxAlignedToLongestEdge(oSurfaceContourLines, oContourLines, oBasePlane, eAdjustmentType, eDistanceMeasureType, eOriginPos)

            Dim oAdjustedCorners As IJElements
            Set oAdjustedCorners = oJigFloorUtil.GetAdjustedPinBedCorners(oCorners, oContourLines, oBasePlane, eAdjustmentType, dTransOffset, dLongiOffset, dTransMinOverHang, dLongiMinOverHang, dAddRow, dAddCol, dTransIntvl, dLongiIntvl)

            Set IJDPinJigAlignmentRule_GetAdjustedPinBedPositions = oJigFloorUtil.OrderPinBedCorners(oContourLines, oAdjustedCorners, oBasePlane, eOriginPos)
             
        Case AlignToLeftContour, AlignToBottomContour, AlignToRightContour, AlignToTopContour
            
            eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_ContourCorners Or STRMFG_PinJig_Adjustment_Offset_Fixed Or STRMFG_PinJig_Adjustment_Extension_Yes
            eOriginPos = oPinJig.PinOriginPosition
            eDistanceMeasureType = ProjectedJigStraightLength
         
            Set oSurfaceContourLines = GetPinJigContourLines(oPinJig, JigSurfaceRemarking)
            Set oContourLines = GetPinJigContourLines(oPinJig, ProjectedJigRemarking)
            
            Dim eMfgJigContourDirectionType As MfgJigContourDirectionType
            Select Case lPinJigAlignmentType
                Case AlignToLeftContour
                    eMfgJigContourDirectionType = Aft
                Case AlignToBottomContour
                    eMfgJigContourDirectionType = Lower
                Case AlignToRightContour
                    eMfgJigContourDirectionType = Fore
                Case AlignToTopContour
                    eMfgJigContourDirectionType = Upper
            End Select
            
            
            Set oCorners = oJigFloorUtil.GetBoxAlignedToInputContour(oSurfaceContourLines, oContourLines, oBasePlane, eAdjustmentType, eDistanceMeasureType, eOriginPos, eMfgJigContourDirectionType)
            Set oOrderedCorners = oJigFloorUtil.OrderPinBedCorners(oContourLines, oCorners, oBasePlane, eOriginPos)
            Set IJDPinJigAlignmentRule_GetAdjustedPinBedPositions = oJigFloorUtil.GetAdjustedPinBedCorners(oOrderedCorners, oContourLines, oBasePlane, eAdjustmentType, dTransOffset, dLongiOffset, dTransMinOverHang, dLongiMinOverHang, dAddRow, dAddCol, dTransIntvl, dLongiIntvl)
                       
        Case AlignToGlobal
        
            'NOTE: STRMFG_PinJig_Adjustment_BoxPoints_Corners is NOT APPLICABLE for this alignment option
            eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_Contours Or STRMFG_PinJig_Adjustment_Offset_Average
            eOriginPos = oPinJig.PinOriginPosition
            
            Set oContourLines = GetPinJigContourLines(oPinJig, ProjectedJigRemarking)
            Set oCorners = oJigFloorUtil.GetRangeBoxOfProjectedContours(oContourLines, oBasePlane)
            
            Set oOrderedCorners = oJigFloorUtil.OrderPinBedCorners(oContourLines, oCorners, oBasePlane, eOriginPos)
            
            Set IJDPinJigAlignmentRule_GetAdjustedPinBedPositions = oJigFloorUtil.GetAdjustedPinBedCorners(oOrderedCorners, oContourLines, oBasePlane, eAdjustmentType, dTransOffset, dLongiOffset, dTransMinOverHang, dLongiMinOverHang, dAddRow, dAddCol, dTransIntvl, dLongiIntvl)
            
        Case AlignToPinsOnCenterLine
            
            eAdjustmentType = STRMFG_PinJig_Adjustment_BoxPoints_Contours Or STRMFG_PinJig_Adjustment_Offset_Average
            eOriginPos = oPinJig.PinOriginPosition
            
            Set oContourLines = GetPinJigContourLines(oPinJig, ProjectedJigRemarking)
            Set oCorners = oJigFloorUtil.GetBoxAlignedToCenterLine(oContourLines, oBasePlane, oRemarkingSurf, eAdjustmentType)
                        
            Set oOrderedCorners = oJigFloorUtil.OrderPinBedCorners(oContourLines, oCorners, oBasePlane, eOriginPos)
            
            Set IJDPinJigAlignmentRule_GetAdjustedPinBedPositions = oJigFloorUtil.GetAdjustedPinBedCornersBasedOnCenterLine(oOrderedCorners, oContourLines, oBasePlane, oRemarkingSurf, eAdjustmentType, dTransOffset, dLongiOffset, dTransMinOverHang, dLongiMinOverHang, dAddRow, dAddCol, dTransIntvl, dLongiIntvl)
    End Select
       
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5027, , "RULES")
End Function

'The tolerances would be used in the following order:
'Offset would be applied on the PinBed.
'Later, MinOverHang check would be done
'Later additonal row/col would be added

'Properties "TransverseOffsetForOriginPin" and "LongitudinalOffsetForOriginPin" are best explained by below figure
'  NEGATIVE OFFSET         POSITIVE OFFSET
'  x  | x    x             | x    x
'     |                    |
'  x  | x    x             | x    x
'     +-------             +-------
'  x    x    x
' We can apply an offset on the Pin Bed so that we avoid pins along the edge
' or a pin exactly at the Pin Jig's seam origin.
'
' NEGATIVE Offset would move the pin origin outwards from the contours
' POSITIVE Offset would move the pin origin inwards from the contours
'
' These properties allow you to control these offsets along the transverse and longitudinal directions respectively.
'
' These offset values are used ONLY when Pin Jig is created WITHOUT ANY ROTATION/MOVEMENT APPLIED ON JIG FLOOR.
'
' The values returned are FRACTIONS of their respective intervals i.e., between 0.0 and 1.0.
'
' To offset the pin bed by an absolute amount, use "ValForRounding" argument to calculate the ratio.
' (i.e., ValForRounding will contain the corresponding Pin Interval)


' Properties "TransverseMinimumOverhang" and "LongitudinalMinimumOverhang" are best explained by below figure
'
'       A  B  C  D        Property "TransverseMinimumOverhang" specifies the minimum overhang that 'O' should have
'                         beyond column 'C'. (i.e) Minimum overhang that would be respected in the transversal direction.
'      \
'       k  x  x  x        Property "LongitudinalMinimumOverhang" is similar to "TransverseMinimumOverhang"
'        \                except that it specifies the longitudinal minimum overhang.
'       x \x  x  x
'          \              The values returned are FRACTIONS of their respective intervals i.e., between 0.0 and 1.0.
'       x  x\ x  x
'            \
'       x  x  O  x        The input argument "ValForRounding" will contain the pin interval corresponding to "ToleranceContext".
'            /            If ToleranceContext is "TransverseMinimumOverhang", ValForRounding will be Transverse Interval.
'       x  x/ x  x        If ToleranceContext is "LongitudinalMinimumOverhang", ValForRounding will be Longitudinal Interval.
'          /
'       x /x  x  x        E.g., if you desire to have a row ONLY if the overhang exceeds half the
'        /                longitudinal interval then return 0.5 for "LongitudinalMinimumOverhang".
'       k  x  x  x
'      /                  E.g., if you desire to have a column ONLY if the overhang exceeds
'                         100 mm (assume DBU is meters), then return (0.1 / ValForRounding)
'       A  B  C  D        for "TransverseMinimumOverhang". (i.e) a minimum of 100mm overhang would be maintained in the transversal direction
'


'Properties "AddAdditionalRow" and "AddAdditionalColumn" indicate whether to add an additional row, column.
'To always have an additional row or column (e.g., like a buffer zone), return ONE.
'To never have an additional row or column, return ZERO.

Private Function IJDPinJigAlignmentRule_SpecifyValueForPurpose(ByVal ValueContext As StrMfgPinJigValueTypes, ByVal pPinJig As Object, Optional ByVal ValForRounding As Double = 0#) As Variant
    Const METHOD = "IJDPinJigAlignmentRule_SpecifyValueForPurpose"
    On Error GoTo ErrorHandler
    
    Dim Retvalue As Variant
    Retvalue = GetValueBasedOnContext(ValueContext, pPinJig, ValForRounding)
        
    IJDPinJigAlignmentRule_SpecifyValueForPurpose = Retvalue
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5026, , "RULES")
End Function

Private Function GetValueBasedOnContext(ByVal ValueContext As StrMfgPinJigValueTypes, _
                                            ByVal pPinJig As Object, _
                                            Optional ByVal ValForRounding As Double = 0#) As Variant
                                       
        Const METHOD = "GetValueBasedOnContext"
        On Error GoTo ErrorHandler
                
        Dim oPinJig As IJPinJig
        Set oPinJig = pPinJig
                
        'Get the pin alignment type.
        Dim lPinJigAlignmentType As Long
        lPinJigAlignmentType = oPinJig.PinJigAlignmentType
        
        Dim Retvalue As Variant
        Select Case lPinJigAlignmentType
    
            Case AlignToDefault
                Select Case ValueContext
                    Case TransverseOffsetForOriginPin
                        Retvalue = -0.5 'Move outwards
                    Case LongitudinalOffsetForOriginPin
                        Retvalue = -0.5
                    Case TransverseMinimumOverhang
                        Retvalue = 0
                    Case LongitudinalMinimumOverhang
                        Retvalue = 0
                    Case AddAdditionalRow
                        Retvalue = 1 'add
                    Case AddAdditionalColumn
                        Retvalue = 1 'add
                    Case PinOriginPosition
                        Retvalue = LowerLeftCorner
                    Case ViewUpVector
                        Retvalue = StrMfgPinJigViewVectorTypes.LowerToUpper
                    Case LocalCoordinateSystemOrigin
                        Retvalue = LowerLeftCorner
                    Case RoundToDegreeMultiple
                        Retvalue = 0
                End Select
            Case AlignToLongestEdge
                Select Case ValueContext
                    Case TransverseOffsetForOriginPin
                        Retvalue = 0.1 / ValForRounding  'Move 100mm inwards
                    Case LongitudinalOffsetForOriginPin
                        Retvalue = 0.1 / ValForRounding
                    Case TransverseMinimumOverhang
                        Retvalue = 0.1 / ValForRounding
                    Case LongitudinalMinimumOverhang
                        Retvalue = 0.1 / ValForRounding
                     Case AddAdditionalRow
                        Retvalue = 0 'donot add
                    Case AddAdditionalColumn
                        Retvalue = 0 'donot add
                    Case PinOriginPosition
                        Retvalue = UpperLeftCorner
                    Case ViewUpVector
                        Retvalue = StrMfgPinJigViewVectorTypes.LowerToUpper
                    Case LocalCoordinateSystemOrigin
                        Retvalue = LowerLeftCorner
                    Case RoundToDegreeMultiple
                        Retvalue = 0
                End Select
                
            Case AlignToLeftContour, AlignToBottomContour, AlignToRightContour, AlignToTopContour
            
                Select Case ValueContext
                    Case TransverseOffsetForOriginPin
                        Retvalue = -0.3 / ValForRounding 'Move 300mm outnwards
                    Case LongitudinalOffsetForOriginPin
                        Retvalue = -0.2 / ValForRounding 'Move 200mm outnwards
                    Case TransverseMinimumOverhang
                        Retvalue = 0.1 / ValForRounding
                    Case LongitudinalMinimumOverhang
                        Retvalue = 0.1 / ValForRounding
                     Case AddAdditionalRow
                        Retvalue = 0 'donot add
                    Case AddAdditionalColumn
                        Retvalue = 0 'donot add
                    Case PinOriginPosition
                        Retvalue = UpperLeftCorner
                    Case ViewUpVector
                        Retvalue = StrMfgPinJigViewVectorTypes.LowerToUpper
                    Case LocalCoordinateSystemOrigin
                        Retvalue = LowerLeftCorner
                    Case RoundToDegreeMultiple
                        Retvalue = 0
                End Select
            Case AlignToGlobal
                 Select Case ValueContext
                   Case TransverseOffsetForOriginPin
                        Retvalue = 0.1 / ValForRounding 'Move 100mm inwards
                    Case LongitudinalOffsetForOriginPin
                        Retvalue = 0.1 / ValForRounding
                    Case TransverseMinimumOverhang
                        Retvalue = 0.1 / ValForRounding
                    Case LongitudinalMinimumOverhang
                        Retvalue = 0.1 / ValForRounding
                     Case AddAdditionalRow
                        Retvalue = 0 'donot add
                    Case AddAdditionalColumn
                        Retvalue = 0 'donot add
                    Case PinOriginPosition
                        Retvalue = UpperLeftCorner
                    Case ViewUpVector
                        Retvalue = StrMfgPinJigViewVectorTypes.LowerToUpper
                    Case LocalCoordinateSystemOrigin
                        Retvalue = LowerLeftCorner
                    Case RoundToDegreeMultiple
                        Retvalue = 0
                 End Select
            Case AlignToPinsOnCenterLine
                 Select Case ValueContext
                   Case TransverseOffsetForOriginPin
                        Retvalue = 0.1 / ValForRounding 'Move 100mm inwards
                    Case LongitudinalOffsetForOriginPin
                        Retvalue = 0.1 / ValForRounding
                    Case TransverseMinimumOverhang
                        Retvalue = 0.1 / ValForRounding
                    Case LongitudinalMinimumOverhang
                        Retvalue = 0.1 / ValForRounding
                     Case AddAdditionalRow
                        Retvalue = 0 'donot add
                    Case AddAdditionalColumn
                        Retvalue = 0 'donot add
                    Case PinOriginPosition
                        Retvalue = UpperLeftCorner
                    Case ViewUpVector
                        Retvalue = StrMfgPinJigViewVectorTypes.LowerToUpper
                    Case LocalCoordinateSystemOrigin
                        Retvalue = LowerLeftCorner
                    Case RoundToDegreeMultiple
                        Retvalue = 0
                 End Select
            'This is the case where move/rotate has been applied on the PinJig or in modification of the PinJig
            'Minimum Overhang values can be specified after move/rotate. Offset values are not applicable
            Case -1
                Select Case ValueContext
                    Case TransverseMinimumOverhang
                        Retvalue = 0.1 / ValForRounding
                    Case LongitudinalMinimumOverhang
                        Retvalue = 0.1 / ValForRounding
                    Case AddAdditionalRow
                        Retvalue = 0 'donot add
                    Case AddAdditionalColumn
                        Retvalue = 0 'donot add
                    Case ViewUpVector
                        Retvalue = StrMfgPinJigViewVectorTypes.LowerToUpper
                    Case LocalCoordinateSystemOrigin
                        Retvalue = LowerLeftCorner
                    Case RoundToDegreeMultiple
                        Retvalue = 0
                End Select
        End Select
        
        GetValueBasedOnContext = Retvalue
    
        Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5026, , "RULES")
End Function


