VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigPartSurfLCRule
' Module: CustomRule
'
' Description: This rule set uses the part surface as the remarking surface and
'              logical connections for the remarking geometries
'
'*******************************************************************************

Option Explicit
Private Const MODULE As String = "MfgPinJigPartSurfLCRule.CustomRule"

Implements IJDPinJigRemarkingRule
Implements IJDPinJigOutputNamingRule


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'METHOD GetConnectedPartsName
'Logic:
'   Get the parts from the system
'   Get the names of the parts
'   Check if the parts intersect the remarking surface
'   if they intersect, return the part names
'if more than on epart intersects the remarking surface, concatenate the part names and return as output
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetConnectedPartsName(RemarkLineObj As Object, oSystem As IJSystem) As String
    Const METHOD = "GetConnectedPartsName"
    On Error GoTo ErrorHandler

    Dim oConnectedParts As IJElements
    Set oConnectedParts = New JObjectCollection

    'From remarkingline get the pinjig
    Dim oGeomchild As IJMfgGeomChild
    Set oGeomchild = RemarkLineObj

    Dim oJigProcessData As IJJigProcessData
    Set oJigProcessData = oGeomchild.GetParent

    Set oGeomchild = Nothing
    Set oGeomchild = oJigProcessData

    Dim oJigPart3d As IJJigPart3D
    Set oJigPart3d = oGeomchild.GetParent

    Set oGeomchild = Nothing
    Set oGeomchild = oJigPart3d

    Dim oPinJig As IJPinJig
    Set oPinJig = oGeomchild.GetParent

    'get the remarking surface
    Dim oRemarkingSurface As IJSurfaceBody
    Set oRemarkingSurface = oPinJig.RemarkingSurface

    'get the plate parts
    Dim oStructDetailHelper As StructDetailHelper
    Dim oEnumPartsUnk       As IEnumUnknown
    Dim oCollectionOfParts  As Collection
    Dim ConvertUtils        As CCollectionConversions
    Dim oParts              As IJElements

    Set oStructDetailHelper = New StructDetailHelper
    oStructDetailHelper.GetPartsDerivedFromSystem oSystem, oEnumPartsUnk, True

    If Not oEnumPartsUnk Is Nothing Then
        'Convert the IEnumUnknown to a VB collection that we can use in VB
        Set ConvertUtils = New CCollectionConversions
        ConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumPartsUnk, oCollectionOfParts
        Set oParts = ConvertUtils.CreateIJElementsCollectionFromVBCollection(oCollectionOfParts)
    End If

'    Check if the plate part is overlapping the remarking surface ( minimum test., intersecting, overlapping )
    
    Dim i                       As Integer
    Dim oObject                 As Object
    Dim strRemarkingName        As String
    Dim dMinDist                As Double
 
    Dim oModelBody1             As IJDModelBody
    Dim oModelBody2             As IJDModelBody
    Dim oClosestPos1            As IJDPosition
    Dim oClosestPos2            As IJDPosition
    
    Set oModelBody1 = oRemarkingSurface
    
    For i = 1 To oParts.Count

        Set oObject = oParts.Item(i)
        
        Set oModelBody2 = oParts.Item(i)
        oModelBody1.GetMinimumDistance oModelBody2, oClosestPos1, oClosestPos2, dMinDist
        
        If dMinDist < 0.01 Then
            oConnectedParts.Add oObject
        End If

    Next i

    Dim oPart                   As Object

    For Each oPart In oConnectedParts
        AppendName oPart, strRemarkingName
    Next

    If oConnectedParts.Count = 0 Then
        AppendName oSystem, strRemarkingName
    End If
    
    GetConnectedPartsName = strRemarkingName

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, "Failed to Get Connected Parts")
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'METHOD GetPartsName
'Logic:
'   Get the parts from the system
'   Get the names of the parts
'   Get the last separator '-' position
'   Compare all the part names till the last '-'
'   If names are different, concatenate all parts names
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetPartsName(RemarkLineObj As Object, oSystem As IJSystem) As String
    Const METHOD = "GetPartsName"
    On Error GoTo ErrorHandler

    'From remarkingline get the pinjig
    Dim oGeomchild As IJMfgGeomChild
    Set oGeomchild = RemarkLineObj

    Dim oJigProcessData As IJJigProcessData
    Set oJigProcessData = oGeomchild.GetParent

    Set oGeomchild = Nothing
    Set oGeomchild = oJigProcessData

    Dim oJigPart3d As IJJigPart3D
    Set oJigPart3d = oGeomchild.GetParent

    Set oGeomchild = Nothing
    Set oGeomchild = oJigPart3d

    Dim oPinJig As IJPinJig
    Set oPinJig = oGeomchild.GetParent

    'get the plate parts
    Dim oStructDetailHelper As StructDetailHelper
    Dim oEnumPartsUnk       As IEnumUnknown
    Dim oCollectionOfParts  As Collection
    Dim ConvertUtils        As CCollectionConversions
    Dim oParts              As IJElements

    Set oStructDetailHelper = New StructDetailHelper
    oStructDetailHelper.GetPartsDerivedFromSystem oSystem, oEnumPartsUnk, True

    If Not oEnumPartsUnk Is Nothing Then
        'Convert the IEnumUnknown to a VB collection that we can use in VB
        Set ConvertUtils = New CCollectionConversions
        ConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumPartsUnk, oCollectionOfParts
        Set oParts = ConvertUtils.CreateIJElementsCollectionFromVBCollection(oCollectionOfParts)

    End If

    Dim i                       As Integer
    Dim j                       As Integer
    Dim oNamedItem              As IJNamedItem

    Set oNamedItem = oParts.Item(1)

    Dim strPart1Name As String
    strPart1Name = oNamedItem.Name

    Dim oPos        As Long

    If oParts.Count > 1 Then
        oPos = InStrRev(strPart1Name, "-", , vbTextCompare)
        Dim strCommon As String

        strCommon = Left(strPart1Name, oPos - 1)
        For i = 2 To oParts.Count
            Set oNamedItem = oParts.Item(i)
            If InStr(1, oNamedItem.Name, strCommon, vbTextCompare) = 0 Then
                strCommon = vbNullString
                Exit For
            End If
        Next i

        If strCommon = vbNullString Then
            'concatenate all string names and return that as a string
            For j = 1 To oParts.Count
                Set oNamedItem = oParts.Item(j)
                If Len(strCommon) > 0 Then strCommon = strCommon & ", "

                strCommon = strCommon & oNamedItem.Name

            Next
        End If
    Else
        strCommon = strPart1Name
    End If

    GetPartsName = strCommon

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, "Failed to Part Name")
End Function

Private Sub AppendName(oPart As Object, RemarkingName As String, Optional strSeparator As String = ", ")
    Const METHOD = "AppendName"
    On Error GoTo ErrorHandler

    Dim oNamedItem As IJNamedItem
    Set oNamedItem = oPart
    If Not oNamedItem Is Nothing Then
        If Len(RemarkingName) > 0 Then RemarkingName = RemarkingName & strSeparator
        RemarkingName = RemarkingName & oNamedItem.Name
    End If

    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, "Failed in AppendName method")
End Sub

Private Function IJDPinJigOutputNamingRule_RemarkingLine(ByVal RemarkLineObj As Object, _
                                                         ByVal SharedRepEntities As IJElements, _
                                                         ByVal ConsumedRepEntities As IJElements, _
                                                         ByVal LongestRepEntity As Object) As String

    Const METHOD = "IJDPinJigOutputNamingRule_RemarkingLine"
    On Error GoTo ErrorHandler

    Dim RemarkingName           As String
    Dim RemarkingNameConsumed   As String
    Dim oRepresentedEntity      As Object
    Dim oMarkingLine As IJMfgMarkingLines_AE
    Dim oMLObject As Object
    
    If LongestRepEntity Is Nothing Or SharedRepEntities Is Nothing Then Exit Function
    If SharedRepEntities.Count = 0 Then Exit Function

    For Each oRepresentedEntity In SharedRepEntities
        Dim strPartName As String
        If TypeOf oRepresentedEntity Is IJPlateSystem Or TypeOf oRepresentedEntity Is IJStiffenerSystem Then
            strPartName = GetConnectedPartsName(RemarkLineObj, oRepresentedEntity)

            ''SHI Specific - UNCOMMENT BELOW LINE IF YOU DO NOT WANT TO GET PART NAMES BASED ON INTERSECTION
            'strPartName = GetPartsName(RemarkLineObj, oRepresentedEntity)

            If strPartName <> vbNullString Then
                If Len(RemarkingName) > 0 Then RemarkingName = RemarkingName & ", "
                RemarkingName = RemarkingName & strPartName
            End If
        ElseIf TypeOf oRepresentedEntity Is IJMfgMarkingLines_AE Then
            Set oMarkingLine = oRepresentedEntity
            Set oMLObject = oMarkingLine.GetMfgMarkingRelatedObject
            AppendName oMLObject, RemarkingName
            Set oMarkingLine = Nothing
            Set oMLObject = Nothing
        Else
            AppendName oRepresentedEntity, RemarkingName
        End If
    Next

    If Not ConsumedRepEntities Is Nothing Then
        If ConsumedRepEntities.Count > 0 Then
            RemarkingName = RemarkingName & ".  ("
            For Each oRepresentedEntity In ConsumedRepEntities
                If TypeOf oRepresentedEntity Is IJPlateSystem Or TypeOf oRepresentedEntity Is IJStiffenerSystem Then
                    strPartName = GetConnectedPartsName(RemarkLineObj, oRepresentedEntity)

                    ''SHI Specific - UNCOMMENT BELOW LINE IF YOU DO NOT WANT TO GET PART NAMES BASED ON INTERSECTION
                    'strPartName = GetPartsName(RemarkLineObj, oRepresentedEntity)

                    If strPartName <> vbNullString Then
                        If Len(RemarkingNameConsumed) > 0 Then RemarkingNameConsumed = RemarkingNameConsumed & ", "
                        RemarkingNameConsumed = RemarkingNameConsumed & strPartName
                    End If
                ElseIf TypeOf oRepresentedEntity Is IJMfgMarkingLines_AE Then
                    Set oMarkingLine = oRepresentedEntity
                    Set oMLObject = oMarkingLine.GetMfgMarkingRelatedObject
                    AppendName oMLObject, RemarkingNameConsumed
                    Set oMarkingLine = Nothing
                    Set oMLObject = Nothing
                Else
                    AppendName oRepresentedEntity, RemarkingNameConsumed
                End If
            Next
            RemarkingName = RemarkingName & RemarkingNameConsumed & ")."
        End If
    End If

    IJDPinJigOutputNamingRule_RemarkingLine = RemarkingName
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5022, , "RULES")
End Function

Private Function IJDPinJigRemarkingRule_GetLocateFilterForPurpose(ByVal PartialPinJig As Object, ByVal AttributeName As String, ByVal AttributeValue As Variant, ByVal RemarkingElems As IJElements) As String
    Const METHOD = "IJDPinJigRemarkingRule_GetLocateFilterForPurpose"
    On Error GoTo ErrorHandler

    Dim oPinJigRule As IJDMfgPinJigRulePerRemarkingType
    Dim oRuleObj    As Object
    'Set IJDPinJigRemarkingRule_GetLocateFilterForPurpose = Nothing

    Select Case AttributeName
        Case "SeamRemark"
            Set oPinJigRule = New SeamRemarking
        Case "PlateRemark"
            Set oPinJigRule = New PlateRemark
        Case "ProfileRemark"
            Set oPinJigRule = New ProfileRemark
        Case "GridLineX_Remark"
            Set oPinJigRule = New FrameLines
        Case "GridLineY_Remark"
            Set oPinJigRule = New Longitudinal
        Case "GridLineZ_Remark"
            Set oPinJigRule = New ButtockLines
        Case "RefCurveRemark"
            Set oPinJigRule = New NavalArchLines
        Case "UserRemark"
            Set oPinJigRule = New UserMarks
        Case "UserExtend"
            Set oPinJigRule = New ExtendUserMarks
        Case "SeamControlRemark"  ' Newly added
            Set oPinJigRule = New SeamControlMark
        Case Default
            Exit Function
    End Select

    If Not oPinJigRule Is Nothing Then
        IJDPinJigRemarkingRule_GetLocateFilterForPurpose = oPinJigRule.GetLocateFilterForRemarkingEntities(PartialPinJig, RemarkingElems)
        Set oPinJigRule = Nothing
    End If

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5022, , "RULES")
End Function

Private Function IJDPinJigRemarkingRule_GetRemarkingEntitiesForPurpose(ByVal PartialPinJig As Object, ByVal AttributeName As String, ByVal AttributeValue As Variant) As IJElements
    Const METHOD = "IJDPinJigRemarkingRule_GetRemarkingEntitiesForPurpose"
    On Error GoTo ErrorHandler

    Dim oPinJigRule As IJDMfgPinJigRulePerRemarkingType
    Dim oRuleObj    As Object
    
    Set IJDPinJigRemarkingRule_GetRemarkingEntitiesForPurpose = Nothing

    Select Case AttributeName
        Case "SeamRemark"
            Set oPinJigRule = New SeamRemarking
        Case "PlateRemark"
            Set oPinJigRule = New PlateRemark
        Case "ProfileRemark"
            Set oPinJigRule = New ProfileRemark
        Case "GridLineX_Remark"
            Set oPinJigRule = New FrameLines
        Case "GridLineY_Remark"
            Set oPinJigRule = New Longitudinal
        Case "GridLineZ_Remark"
            Set oPinJigRule = New ButtockLines
        Case "RefCurveRemark"
            Set oPinJigRule = New NavalArchLines
        Case "UserRemark"
            Set oPinJigRule = New UserMarks
        Case "UserExtend"
            Set oPinJigRule = New ExtendUserMarks
        Case Default
            Exit Function
    End Select

    If Not oPinJigRule Is Nothing Then
        Set IJDPinJigRemarkingRule_GetRemarkingEntitiesForPurpose = oPinJigRule.GetEntitiesForRemarking(PartialPinJig)
        Set oPinJigRule = Nothing
    End If

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5022, , "RULES")
End Function

'---------------------------------------------------------------------------------------
' Procedure : IJDPinJigRemarkingRule_GetRemarkingGeometryForPurpose
' Purpose   : Return remarking geometry corresponding to the "AttributeName"
'---------------------------------------------------------------------------------------
Private Function IJDPinJigRemarkingRule_GetRemarkingGeometryForPurpose(ByVal PartialPinJig As Object, ByVal AttributeName As String, ByVal AttributeValue As Variant, ByVal RemarkingElems As IJElements) As IJMfgGeomCol3d
 Const METHOD = "IJDPinJigRemarkingRule_GetRemarkingGeometryForPurpose"
    On Error GoTo ErrorHandler

    Dim oPinJigRule As IJDMfgPinJigRulePerRemarkingType
    Dim oRuleObj    As Object

    Select Case AttributeName
        Case "SeamRemark"
            Set oPinJigRule = New SeamRemarking
        Case "PlateRemark"
             Select Case AttributeValue
                Case 1
                    Set oPinJigRule = New PlateRemark
                Case 2
                   Set oPinJigRule = New PlateRMLBoth
                Case 3
                   Set oPinJigRule = New PlateRMLBothnLC
            End Select
        Case "ProfileRemark"
            Set oPinJigRule = New ProfileRemark
        Case "GridLineX_Remark"
            Set oPinJigRule = New FrameLines
        Case "GridLineY_Remark"
            Set oPinJigRule = New Longitudinal
        Case "GridLineZ_Remark"
            Set oPinJigRule = New ButtockLines
        Case "RefCurveRemark"
            Set oPinJigRule = New NavalArchLines
        Case "UserRemark"
            Set oPinJigRule = New UserMarks
        Case "UserExtend"
            Set oPinJigRule = New ExtendUserMarks
        Case Default
            Exit Function
    End Select

    Set IJDPinJigRemarkingRule_GetRemarkingGeometryForPurpose = oPinJigRule.GetRemarkingGeometry(PartialPinJig, RemarkingElems)

    Set oPinJigRule = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5022, , "RULES")
End Function

Private Function IJDPinJigRemarkingRule_GetRemarkingSurface(ByVal PartialPinJig As Object) As IJSurfaceBody
    Const METHOD As String = "GetRemarkingSurface"
    On Error GoTo ErrorHandler
    
    'Check if there are any gaps between the part surfaces.
    'If No gaps
        'Get the surfaces stitched on the remarking sides
    'If there are gaps
        'Get the offset value from the plate system (weighted thickness average)
        'Get surface based on the offset from the molded sides of the plate part
    
    'To check for gaps:
    'For every supported plate part get the connected supported plate parts(surrounding plates)
    'Find minimum distance between these plate part surfaces
    'If less than Tolerance, means NO gaps
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
    
    Dim RootX As Double, RootY As Double, RootZ As Double
    Dim NormX As Double, NormY As Double, NormZ As Double
    oPinJig.GetBasePlane NormX, NormY, NormZ, RootX, RootY, RootZ
    
    Dim oPlane As IJPlane
    Set oPlane = New Plane3d

    oPlane.DefineByPointNormal RootX, RootY, RootZ, NormX, NormY, NormZ
    
    Dim oSurfaceUtil As IJMfgUtilSurface
    Set oSurfaceUtil = New MfgUtilSurface
    
    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates
    
    If AreThePlatePartSurfacesHavingGaps(PartialPinJig) = True Then 'There are gaps.
        
        Dim RemarkSides() As Long
        RemarkSides = oSurfaceUtil.GetRemarkingSidesOfPlates(oPlateColl, oPlane)
        
        Dim dRemarkingSurfaceOffset As Double
        dRemarkingSurfaceOffset = oSurfaceUtil.GetOffsetValueFromPlateSystem(oPlateColl, RemarkSides, SurfaceOffsetType_WeightedThicknessAverage)
    
        'Get the surface based on offset from molded side of the plate part
        Set IJDPinJigRemarkingRule_GetRemarkingSurface = _
            oSurfaceUtil.GetSurfaceBasedOnOffsetFromPlatePart(oPlateColl, RemarkSides, SurfaceOffsetType_InputArgument, dRemarkingSurfaceOffset)
        
    Else 'No gaps
        
        Set IJDPinJigRemarkingRule_GetRemarkingSurface = _
            oSurfaceUtil.GenSurfFromOppositeSidesFacingPlane(oPlateColl, oPlane)
   
    End If
    
    Set oPlane = Nothing
    Set oSurfaceUtil = Nothing
    Set oPinJig = Nothing
    Set oPlateColl = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, "Failed to get remarking surface")
End Function

Private Function IJDPinJigRemarkingRule_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object, ByVal AttributeName As String, ByVal AttributeValue As Variant) As Long()

End Function

Private Sub IJDPinJigRemarkingRule_PostProcess3dRemarkingLines(ByVal PartialPinJig As Object, AllRemarkGeom As IJMfgGeomCol3d)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : IJDPinJigRemarkingRule_SpecifyRemarkingTypesForPurpose
' Purpose   : For the "Purpose" specified by the input string, return a list of
'             remarking types that satisfy that particular purpose.
'
' General convention for the "Purpose" string:
'    Prefixed with "IncludeIn_":   Only types specified here will be included.
'    Prefixed with "ExcludeFrom_": Except for types specified here, all others will be included.
'
'---------------------------------------------------------------------------------------

Private Function IJDPinJigRemarkingRule_SpecifyRemarkingTypesForPurpose(ByVal PurposeOfRemarkingTypes As String) As Long()
    Const METHOD = "IJDPinJigRemarkingRule_SpecifyRemarkingTypesForPurpose"
    On Error GoTo ErrorHandler

    Dim RemarkTypes() As Long
    Select Case PurposeOfRemarkingTypes

        Case "IncludeIn_MarkingCommand"
            'Specify the types of remarking lines that should be displayed
            'within the RAD 2D environment of the marking command.
            ReDim RemarkTypes(1 To 2) As Long
            RemarkTypes(1) = STRMFG_PinJig_Remarking_Plate
            RemarkTypes(2) = STRMFG_PinJig_Remarking_Profile

        Case "ExcludeFrom_IntersectionPointCreation"
            'Specify the types of remarking lines that should be excluded
            'from participating in the intersection point creation process.
            ReDim RemarkTypes(1 To 3) As Long
            RemarkTypes(1) = STRMFG_NAVALARCHLINE
            RemarkTypes(2) = STRMFG_PinJig_Remarking_NavalArch
            RemarkTypes(3) = STRMFG_PINJIG_DIAGONAL

        Case "RemarkingLinePriority"
            'Specify rank (order of descending priority) of remarking line types.
            'When two remarking lines are identical (they completely overlap each other)
            'the line whose type has higher rank (lower index) will be used for
            'intersection point creation.
            ReDim RemarkTypes(1 To 3) As Long
            RemarkTypes(1) = STRMFG_PinJigContourLine2D ' NB: Should specify 2D, not 3D.
            RemarkTypes(2) = STRMFG_PinJig_Remarking_Frame
            RemarkTypes(3) = STRMFG_PinJig_Remarking_Seam
            'Above example, Contour lines will be preferred over Frames, and
            'frames will be preferred over seams.  Types not listed above will
            '"lose" to types listed above.  If two lines overlap, and neither of
            'their types are listed above, one will be arbitrarily picked.


        Case Else
            ReDim RemarkTypes(0 To 0) As Long

    End Select

    IJDPinJigRemarkingRule_SpecifyRemarkingTypesForPurpose = RemarkTypes

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5023, , "RULES")
End Function

