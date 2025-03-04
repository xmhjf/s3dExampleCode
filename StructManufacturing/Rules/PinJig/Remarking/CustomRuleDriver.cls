VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomRuleDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: CustomRuleDriver
'
' Description:
'
' Author: Anand Hariharan
'
'*******************************************************************************

Option Explicit
Private Const MODULE As String = "MfgPinJigRule.CustomRuleDriver"

Implements IJDPinJigRemarkingRule
Implements IJDPinJigOutputNamingRule

Private Function IJDPinJigOutputNamingRule_RemarkingLine(ByVal RemarkLineObj As Object, _
                                                         ByVal SharedRepEntities As IJElements, _
                                                         ByVal ConsumedRepEntities As IJElements, _
                                                         ByVal LongestRepEntity As Object) As String
    
    Const METHOD = "IJDPinJigOutputNamingRule_RemarkingLine"
    On Error GoTo ErrorHandler

    If LongestRepEntity Is Nothing Or SharedRepEntities Is Nothing Then Exit Function
    If SharedRepEntities.Count = 0 Then Exit Function
    
    Dim RemarkingName As String
    Dim oNamedItem As IJNamedItem
        
    Dim oRepresentedEntity As Object
    
    For Each oRepresentedEntity In SharedRepEntities
        Set oNamedItem = oRepresentedEntity
        If Not oNamedItem Is Nothing Then
            If Len(RemarkingName) > 0 Then RemarkingName = RemarkingName & ", "
            RemarkingName = RemarkingName & oNamedItem.Name
        End If
    Next
    
    Dim iCounter As Integer
    iCounter = 0
    
    If Not ConsumedRepEntities Is Nothing Then
      
        For Each oRepresentedEntity In ConsumedRepEntities
        
            iCounter = iCounter + 1
            If iCounter = 1 Then RemarkingName = RemarkingName & ".  (Also "
            
            Set oNamedItem = oRepresentedEntity
            If Not oNamedItem Is Nothing Then
                RemarkingName = RemarkingName & ", " & oNamedItem.Name
            End If
        Next
        If iCounter > 0 Then RemarkingName = RemarkingName & ")."
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
    'Set IJDPinJigRemarkingRule_GetLocateFilterForPurpose = Nothing

    Select Case AttributeName
        Case "SeamRemark"
            Set oPinJigRule = New SeamRemark_ButtPC
        Case "PlateRemark"
            Set oPinJigRule = New PhyConnPlateRemarking
        Case "ProfileRemark"
            Set oPinJigRule = New PhyConnProfileRemarking
        Case "GridLineX_Remark"
            Set oPinJigRule = New FrameLines
        Case "GridLineY_Remark"
            Set oPinJigRule = New LongitudinalLines
        Case "GridLineZ_Remark"
            Set oPinJigRule = New ButtockLines
        Case "RefCurveRemark"
            Set oPinJigRule = New NavalArchLines
        Case "UserRemark"
            Set oPinJigRule = New CreateLinesFromUserMarks
        Case "UserExtend"
            Set oPinJigRule = New ExtendLinesFromUserMarks
        Case "SeamControlRemark"  ' Newly added
            Set oPinJigRule = New SeamControlRemark
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
    Set IJDPinJigRemarkingRule_GetRemarkingEntitiesForPurpose = Nothing

    Select Case AttributeName
        Case "SeamRemark"
            Set oPinJigRule = New SeamRemark_ButtPC
        Case "PlateRemark"
            Set oPinJigRule = New PhyConnPlateRemarking
        Case "ProfileRemark"
            Set oPinJigRule = New PhyConnProfileRemarking
        Case "GridLineX_Remark"
            Set oPinJigRule = New FrameLines
        Case "GridLineY_Remark"
            Set oPinJigRule = New LongitudinalLines
        Case "GridLineZ_Remark"
            Set oPinJigRule = New ButtockLines
        Case "RefCurveRemark"
            Set oPinJigRule = New NavalArchLines
        Case "UserRemark"
            Set oPinJigRule = New CreateLinesFromUserMarks
        Case "UserExtend"
            Set oPinJigRule = New ExtendLinesFromUserMarks
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
Private Function IJDPinJigRemarkingRule_GetRemarkingGeometryForPurpose(ByVal PartialPinJig As Object, ByVal AttributeName As String, ByVal AttributeValue As Variant, ByVal RemarkingElems As IMSCoreCollections.IJElements) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
 Const METHOD = "IJDPinJigRemarkingRule_GetRemarkingGeometryForPurpose"
    On Error GoTo ErrorHandler

    Dim oPinJigRule As IJDMfgPinJigRulePerRemarkingType

    Select Case AttributeName
        Case "SeamRemark"
            Set oPinJigRule = New SeamRemark_ButtPC
        Case "PlateRemark"
           Select Case AttributeValue
                Case 1
                    Set oPinJigRule = New PhyConnPlateRemarking
                Case 2
                   Set oPinJigRule = New PCPlateRMLBoth
                Case 3
                   Set oPinJigRule = New PCPlateRMLBothWithLC
            End Select
        Case "ProfileRemark"
            Set oPinJigRule = New PhyConnProfileRemarking
        Case "GridLineX_Remark"
            Set oPinJigRule = New FrameLines
        Case "GridLineY_Remark"
            Set oPinJigRule = New LongitudinalLines
        Case "GridLineZ_Remark"
            Set oPinJigRule = New ButtockLines
        Case "RefCurveRemark"
            Set oPinJigRule = New NavalArchLines
        Case "UserRemark"
            Set oPinJigRule = New CreateLinesFromUserMarks
        Case "UserExtend"
            Set oPinJigRule = New ExtendLinesFromUserMarks
        Case "SeamControlRemark"  ' Newly added
            Set oPinJigRule = New SeamControlRemark
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
    
    Set IJDPinJigRemarkingRule_GetRemarkingSurface = _
        oSurfaceUtil.GenSurfFromOppositeSidesFacingPlane(oPinJig.SupportedPlates, oPlane)
    
    m_dRemarkingSurfaceOffset = 0#

    Set oPlane = Nothing
    Set oSurfaceUtil = Nothing
    Set oPinJig = Nothing
    
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


