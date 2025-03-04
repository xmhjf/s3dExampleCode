Attribute VB_Name = "Evaluates"
'
'File : Evaluates.bas
'
'Author : MH
'
'Description :
'    module to concentrate on evaluation calculations for cans
'
'History:
'   some time in 2009   MH   creation
'   April 30 2009       GG   Split from InLine.cls, End.cls, and StubEnd.cls
'   May 5 2009          RP   Added evaluate code for stubend and end can
'   May 29 2009         EM   Fixed for ID/OD match (TR 165044)
'   June 08 2009        GG   TR#165047: Need L3HullLength and L3CenterlineLength for End Can
'   June 11 2009        GG   DI#166343: Need Minimum Extension Distance for customized Can type
'   June 12 2009        GG   TR#166471: Fixed the chamfer length calculations
'   June 16 2009        GG   TR#166637: Fixed the chamfer length calculations using the mitered edge thickness of the cone.
'                               Chamfer length is greater than zero only if the Can thickness is greater than the adjusted Cone thickness
'   July 9 2009         MH   TR# 167837.   To maintain user-specified slope, Can rule needs to set "effective coneLength" knowing BUCan creates the molded-form surface on the ID.
'   July 23, 2009       MH   163732 update parameters during compute
'   Aug 05, 2009        GG   169274 Fixed the problem that Stub End L2CenterlineLength is not correct in some cases
'   Mar 04, 2010        MH   175947 go TDL if split location is off the member system
'*****************************************************************************************************************
Option Explicit
Const distTol = 0.000001
Const oneDegree = 0.0174532925          ' in radians
Const eightyNineDegrees = 1.55334303    ' in radians

Private Const MODULE = "Evaluates"

Public Sub InLineEvaluate(ByVal MyGC As SP3DGeometricConstruction.IJGeometricConstruction)
    Const METHOD = "InLineEvaluate"
    On Error GoTo ErrorHandler
    ' use the GC as a GCMacro
    Dim iCanRule As ISPSCanRule
    Dim MyGCMacro As IJGeometricConstructionMacro
    Dim oBUCan As ISPSDesignedMember
    Dim crStatus As SPSCanRuleStatus
    
    Set iCanRule = MyGC
    Set MyGCMacro = MyGC
    
    'MsgBox ("Inline Can Rule Evaluate")
    Dim iPlane1 As IJPlane, iPlane2 As IJPlane
    Dim dCanID As Double, dCanOD As Double
    Dim dNborStartID As Double, dNborStartOD As Double
    Dim dNborEndID As Double, dNborEndOD As Double
    Dim eTubeExtension As Integer, eConeExtension As Integer
    Dim dOverLengthStart As Double, dOverLengthEnd As Double
    Dim bIsHullStart As Boolean, bIsHullEnd As Boolean
    
    Dim dChamferLengthStart As Double, dChamferLengthEnd As Double
    Dim dCone1Length As Double, dCone2Length As Double
    Dim dChamferSlope As Double, dFactor As Double
    Dim dMinimumExtensionDistance As Double
    Dim dRoundOffDistance As Double
    
    Dim posL2Hull As IJDPosition, posL2Centerline As IJDPosition
    Dim posL3Hull As IJDPosition, posL3Centerline As IJDPosition
    Dim oPrimaryMS As ISPSMemberSystem
    Dim vecPrimaryX As Double, vecPrimaryY As Double, vecPrimaryZ As Double
    Dim oAttrColl As Object
    Dim varValue As Variant

    Set oPrimaryMS = MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_Primary).Item("1")
    Set iPlane1 = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_Planes).Item("1")
    Set iPlane2 = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_Planes).Item("2")
    Set oBUCan = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Item("1")

    Dim dCone1Thickness As Double
    Dim dCone2Thickness As Double
    Dim dAdjustedConeThickness As Double, dConeThicknessParameter As Double
    Dim dCanThickness As Double

    ' dDiameterCan is already set on the BUCan by CrossSectionNotify semantic calling UpdateOutputCrossSectionDimensions.
    
    GetDiameters MyGCMacro, SPSMemberAxisAlong, dCanID, dCanOD
    GetDiameters MyGCMacro, SPSMemberAxisStart, dNborStartID, dNborStartOD
    GetDiameters MyGCMacro, SPSMemberAxisEnd, dNborEndID, dNborEndOD
        
    eTubeExtension = 0
    eTubeExtension = MyGC.Parameter(attrL2Method)

    If eTubeExtension = TubeExtension_CLLength Then
        bIsHullStart = False
        dOverLengthStart = MyGC.Parameter(attrL2Length)
        dFactor = dOverLengthStart / dCanOD
        MyGC.Parameter(attrL2Factor) = dFactor

    ElseIf eTubeExtension = TubeExtension_CLFactor Then
        bIsHullStart = False
        dFactor = MyGC.Parameter(attrL2Factor)
        dOverLengthStart = dCanOD * dFactor
        MyGC.Parameter(attrL2Length) = dOverLengthStart

    ElseIf eTubeExtension = TubeExtension_HullLength Then
        bIsHullStart = True
        dOverLengthStart = MyGC.Parameter(attrL2Length)
        dFactor = dOverLengthStart / dCanOD
        MyGC.Parameter(attrL2Factor) = dFactor

    ElseIf eTubeExtension = TubeExtension_HullFactor Then
        bIsHullStart = True
        dFactor = MyGC.Parameter(attrL2Factor)
        dOverLengthStart = dCanOD * dFactor
        MyGC.Parameter(attrL2Length) = dOverLengthStart
    
    Else
        MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "error in start tube extension parameter"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "error in start tube extension parameter"
    End If
    
    eTubeExtension = 0
    eTubeExtension = MyGC.Parameter(attrL3Method)

    If eTubeExtension = TubeExtension_CLLength Then
        bIsHullEnd = False
        dOverLengthEnd = MyGC.Parameter(attrL3Length)
        dFactor = dOverLengthEnd / dCanOD
        MyGC.Parameter(attrL3Factor) = dFactor
    
    ElseIf eTubeExtension = TubeExtension_CLFactor Then
        bIsHullEnd = False
        dFactor = MyGC.Parameter(attrL3Factor)
        dOverLengthEnd = dCanOD * dFactor
        MyGC.Parameter(attrL3Length) = dOverLengthEnd

    ElseIf eTubeExtension = TubeExtension_HullLength Then
        bIsHullEnd = True
        dOverLengthEnd = MyGC.Parameter(attrL3Length)
        dFactor = dOverLengthEnd / dCanOD
        MyGC.Parameter(attrL3Factor) = dFactor

    ElseIf eTubeExtension = TubeExtension_HullFactor Then
        bIsHullEnd = True
        dFactor = MyGC.Parameter(attrL3Factor)
        dOverLengthEnd = dCanOD * dFactor
        MyGC.Parameter(attrL3Length) = dOverLengthEnd
    
    Else
        MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "error in end tube extension parameter"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "error in end tube extension parameter"
    End If

    dChamferSlope = MyGC.Parameter(attrChamferSlope)

'   Calculate Cone1Length and ChamferStartLength
    dCone1Length = 0#
    Dim dConeAngle As Double, dConeSlope As Double
    
''    FindActualPlateThickness MyGC, "Can", dCanThickness       can thickness updated by ISPSCanRuleHelper_UpdateOutputCrossSectionDimensions
    dCanThickness = MyGC.Parameter(attrCanThickness)
    
    If ((dCanID >= dNborStartOD) Or (dCanOD <= dNborStartID)) Then  'Cone1 is required

        FindActualPlateThickness MyGC, "Cone1", dCone1Thickness
        
        dConeThicknessParameter = MyGC.Parameter(attrCone1Thickness)
        If Abs(dConeThicknessParameter - dCone1Thickness) > distTol Then
            MyGC.Parameter(attrCone1Thickness) = dCone1Thickness
        End If
        
        eConeExtension = 0
        eConeExtension = MyGC.Parameter(attrCone1Method)
    
        If eConeExtension = ConeMethod_Length Then
            dCone1Length = Abs(MyGC.Parameter(attrCone1Length))
            
            If dCone1Length < distTol Then
                dConeAngle = 0
                dConeSlope = 0
                dCone1Length = 0
                dAdjustedConeThickness = dCone1Thickness
    
            Else
                ComputeConeAngle dCanOD, dCanThickness, dCone1Length, dNborStartOD, dCone1Thickness, dConeAngle, dAdjustedConeThickness
                dConeSlope = 1# / Tan(dConeAngle)
            
            End If

            MyGC.Parameter(attrCone1Slope) = dConeSlope
            MyGC.Parameter(attrCone1Angle) = dConeAngle
        
        ElseIf eConeExtension = ConeMethod_Slope Then
            dConeSlope = Abs(MyGC.Parameter(attrCone1Slope))

            If dConeSlope > 100 Or dConeSlope < distTol Then
                dCone1Length = 0
                dConeAngle = 0
                dAdjustedConeThickness = dCone1Thickness

            Else
                dConeAngle = Atn(1# / dConeSlope)
                dAdjustedConeThickness = dCone1Thickness / Cos(dConeAngle)
                ComputeConeLength dConeAngle, dNborStartOD, dCanOD, dAdjustedConeThickness, dCanThickness, dCone1Length

            End If

            MyGC.Parameter(attrCone1Length) = dCone1Length
            MyGC.Parameter(attrCone1Angle) = dConeAngle

        ElseIf eConeExtension = ConeMethod_Angle Then
            dConeAngle = Abs(MyGC.Parameter(attrCone1Angle))
            If dConeAngle < oneDegree Or dConeAngle > eightyNineDegrees Then
                dCone1Length = 0
                dConeSlope = 0
                dAdjustedConeThickness = dCone1Thickness

            Else
                dAdjustedConeThickness = dCone1Thickness / Cos(dConeAngle)
                dConeSlope = 1# / Tan(dConeAngle)
                ComputeConeLength dConeAngle, dNborStartOD, dCanOD, dAdjustedConeThickness, dCanThickness, dCone1Length
            End If

            MyGC.Parameter(attrCone1Length) = dCone1Length
            MyGC.Parameter(attrCone1Slope) = dConeSlope

        Else
            MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
            WriteToErrorLog E_FAIL, MODULE, METHOD, "error in cone 1 extension parameter (Length)"
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "error in cone 1 extension parameter (Length)"
        End If
        
        ComputeChamferLengthAtCone dChamferSlope, dCanThickness, dAdjustedConeThickness, dCone1Length, dChamferLengthStart

    Else

        ComputeChamferLengthAtNbor dChamferSlope, dCanOD, dCanID, dNborStartOD, dNborStartID, dChamferLengthStart

        dCone1Length = 0#
        dCone1Thickness = dCanThickness
    End If

'   Calculate Cone2Length and ChamferEndLength
    dCone2Length = 0#
    If ((dCanID >= dNborEndOD) Or (dCanOD <= dNborEndID)) Then  'Cone2 is required
   
        FindActualPlateThickness MyGC, "Cone2", dCone2Thickness
        
        dConeThicknessParameter = MyGC.Parameter(attrCone2Thickness)
        If Abs(dConeThicknessParameter - dCone2Thickness) > distTol Then
            MyGC.Parameter(attrCone2Thickness) = dCone2Thickness
        End If
        
        eConeExtension = 0
        eConeExtension = MyGC.Parameter(attrCone2Method)
    
        If eConeExtension = ConeMethod_Length Then
            dCone2Length = Abs(MyGC.Parameter(attrCone2Length))
            
            If dCone2Length < distTol Then
                dConeAngle = 0
                dConeSlope = 0
                dCone2Length = 0
                dAdjustedConeThickness = dCone2Thickness
    
            Else
                ComputeConeAngle dCanOD, dCanThickness, dCone2Length, dNborEndOD, dCone2Thickness, dConeAngle, dAdjustedConeThickness
                dConeSlope = 1# / Tan(dConeAngle)
            
            End If

            MyGC.Parameter(attrCone2Slope) = dConeSlope
            MyGC.Parameter(attrCone2Angle) = dConeAngle

        ElseIf eConeExtension = ConeMethod_Slope Then
            dConeSlope = Abs(MyGC.Parameter(attrCone2Slope))

            If dConeSlope > 100 Or dConeSlope < distTol Then
                dCone2Length = 0
                dConeAngle = 0
                dAdjustedConeThickness = dCone2Thickness

            Else
                dConeAngle = Atn(1# / dConeSlope)
                dAdjustedConeThickness = dCone2Thickness / Cos(dConeAngle)
                ComputeConeLength dConeAngle, dNborEndOD, dCanOD, dAdjustedConeThickness, dCanThickness, dCone2Length

            End If

            MyGC.Parameter(attrCone2Length) = dCone2Length
            MyGC.Parameter(attrCone2Angle) = dConeAngle
            
        ElseIf eConeExtension = ConeMethod_Angle Then
            dConeAngle = Abs(MyGC.Parameter(attrCone2Angle))
            If dConeAngle < oneDegree Or dConeAngle > eightyNineDegrees Then
                dCone2Length = 0
                dConeSlope = 0
                dAdjustedConeThickness = dCone2Thickness

            Else
                dAdjustedConeThickness = dCone2Thickness / Cos(dConeAngle)
                dConeSlope = 1# / Tan(dConeAngle)
                ComputeConeLength dConeAngle, dNborEndOD, dCanOD, dAdjustedConeThickness, dCanThickness, dCone2Length
            End If

            MyGC.Parameter(attrCone2Length) = dCone2Length
            MyGC.Parameter(attrCone2Slope) = dConeSlope
            
        Else
            MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
            WriteToErrorLog E_FAIL, MODULE, METHOD, "error in cone 2 extension parameter (Length)"
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "error in cone 2 extension parameter (Length)"
        End If
        
        ComputeChamferLengthAtCone dChamferSlope, dCanThickness, dAdjustedConeThickness, dCone2Length, dChamferLengthEnd

    Else

        ComputeChamferLengthAtNbor dChamferSlope, dCanOD, dCanID, dNborEndOD, dNborEndID, dChamferLengthEnd

        dCone2Length = 0#
        dCone2Thickness = dCanThickness
    End If

    GetUnitVectorTangent oPrimaryMS, vecPrimaryX, vecPrimaryY, vecPrimaryZ
    
    iCanRule.Services.ComputeMinMaxPoints 0, posL2Hull, posL2Centerline, posL3Centerline, posL3Hull, crStatus
    If crStatus <> SPSCanRule_Ok Then
        MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "problem computing hull length points"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "problem computing hull length points"
    End If
    
    Dim iIJPoint As IJPoint
    
    'set the position to the middle of the two points returned.
    Set iIJPoint = MyGCMacro
    iIJPoint.SetPoint 0.5 * (posL2Hull.x + posL3Hull.x), 0.5 * (posL2Hull.y + posL3Hull.y), 0.5 * (posL2Hull.z + posL3Hull.z)
    
    'Attributes for IJUASMCanRuleResult
    Dim dCanLength As Double
    Dim dUncutLength As Double
    Dim dL2HullLength As Double
    Dim dL2CenterlineLength As Double
    Dim dL3HullLength As Double
    Dim dL3CenterlineLength As Double
    
    Dim dL2Diff As Double
    Dim dL3Diff As Double
    Dim dRoundOffDiff As Double
    
    dL2Diff = (posL2Centerline.x - posL2Hull.x) * vecPrimaryX + _
                (posL2Centerline.y - posL2Hull.y) * vecPrimaryY + _
                   (posL2Centerline.z - posL2Hull.z) * vecPrimaryZ
    dL3Diff = (posL3Hull.x - posL3Centerline.x) * vecPrimaryX + _
                (posL3Hull.y - posL3Centerline.y) * vecPrimaryY + _
                   (posL3Hull.z - posL3Centerline.z) * vecPrimaryZ
    
    dMinimumExtensionDistance = MinimumExtensionDistance(CanType_InLine, dCanOD, MyGC)
    
    'find Hull lengths first, will set Centerline lengths later, with round off corrections
    If bIsHullStart Then
        dL2HullLength = dOverLengthStart
    Else
        dL2HullLength = dOverLengthStart - dL2Diff
    End If
    If dL2HullLength < dMinimumExtensionDistance Then
    
        MyGC.PostError INVALID_PARAMETERS, False, strCodeListTablename, oBUCan
        ' the codelist simply says the parameter is invalid and refers the user to the error
        ' log, so writh the additional information needed
        WriteToErrorLog S_FALSE, MODULE, METHOD, "The specified " & attrL2Length & " length of " & Format$(dL2HullLength, "#.######") & _
                                                 " does not meet the specified " & _
                                                 " minimum of " & Format$(dMinimumExtensionDistance, "#.######")
           Err.Clear
        dL2HullLength = dMinimumExtensionDistance
    End If
     
    If bIsHullEnd Then
        dL3HullLength = dOverLengthEnd
    Else
        dL3HullLength = dOverLengthEnd - dL3Diff
    End If
    If dL3HullLength < dMinimumExtensionDistance Then
        MyGC.PostError INVALID_PARAMETERS, False, strCodeListTablename, oBUCan
        ' the codelist simply says the parameter is invalid and refers the user to the error
        ' log, so writh the additional information needed
        WriteToErrorLog S_FALSE, MODULE, METHOD, "The specified " & attrL3Length & " length of " & Format$(dL3HullLength, "#.######") & _
                                                 " does not meet the specified " & _
                                                 " minimum of " & Format$(dMinimumExtensionDistance, "#.######")
        dL3HullLength = dMinimumExtensionDistance
    End If
    
    'Uncut length = Can length (Hull lengths = 0) + Hull lengths + cone lengths + chamfer lengths
    dUncutLength = Abs((posL3Hull.x - posL2Hull.x) * vecPrimaryX + (posL3Hull.y - posL2Hull.y) * vecPrimaryY + (posL3Hull.z - posL2Hull.z) * vecPrimaryZ)
    dUncutLength = dUncutLength + dL2HullLength + dL3HullLength + dCone1Length + dChamferLengthStart + dCone2Length + dChamferLengthEnd
    
    'round off uncut length
    dRoundOffDistance = MyGC.Parameter(attrRoundoffDistance)
    Dim dRoundOffUncutLength As Double
    If dRoundOffDistance > dTol Then
        dRoundOffUncutLength = Int((dUncutLength + dRoundOffDistance - dTol) / dRoundOffDistance) * dRoundOffDistance
        dRoundOffDiff = dRoundOffUncutLength - dUncutLength
        dUncutLength = dRoundOffUncutLength
    Else
        dRoundOffUncutLength = dUncutLength
        dRoundOffDiff = 0
    End If
    
    'add round off corrections to hull lengths and centerline lengths
    dL2HullLength = dL2HullLength + dRoundOffDiff / 2
    dL2CenterlineLength = dL2HullLength + dL2Diff
    dL3HullLength = dL3HullLength + dRoundOffDiff / 2
    dL3CenterlineLength = dL3HullLength + dL3Diff
    
    'start plane for uncut can
    posL2Hull.x = posL2Hull.x - (dL2HullLength + dCone1Length + dChamferLengthStart) * vecPrimaryX
    posL2Hull.y = posL2Hull.y - (dL2HullLength + dCone1Length + dChamferLengthStart) * vecPrimaryY
    posL2Hull.z = posL2Hull.z - (dL2HullLength + dCone1Length + dChamferLengthStart) * vecPrimaryZ
    If Not PositionWithinRange(oPrimaryMS, posL2Hull) Then
        MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "hull length points off member system"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "hull length points off member system"
    End If

    UpdatePlane iPlane1, posL2Hull, vecPrimaryX, vecPrimaryY, vecPrimaryZ
    
    'end plane for uncut can
    posL3Hull.x = posL3Hull.x + (dL3HullLength + dCone2Length + dChamferLengthEnd) * vecPrimaryX
    posL3Hull.y = posL3Hull.y + (dL3HullLength + dCone2Length + dChamferLengthEnd) * vecPrimaryY
    posL3Hull.z = posL3Hull.z + (dL3HullLength + dCone2Length + dChamferLengthEnd) * vecPrimaryZ
    If Not PositionWithinRange(oPrimaryMS, posL3Hull) Then
        MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "hull length points off member system"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "hull length points off member system"
    End If

    UpdatePlane iPlane2, posL3Hull, vecPrimaryX, vecPrimaryY, vecPrimaryZ
     
''    If dCanOD < dNborStartOD Then
''       can diameter too small
''
''    ElseIf dCanID > dNborStartID Then
''       bad case of can ID within thickness of nbor
''
''    Else
''       case of chamfer to be applied at end of tube, no cone.
''       chamfer Length = chamferSlope * Max(dCanOD - dNborStartOD, dNborStartID - dCanID)
    
    'can length = length of the uncut can without cones
    dCanLength = dUncutLength - dCone1Length - dCone2Length
    
    'write back attributes for BUCan
    Set oAttrColl = GetAttributeCollection(oBUCan, InterfaceName_IUABuiltUpCan)

    SetAttributeValue oAttrColl, attrDiameterStart, dNborStartOD
    SetAttributeValue oAttrColl, attrLengthStartCone, dCone1Length
    SetAttributeValue oAttrColl, attrDiameterEnd, dNborEndOD
    SetAttributeValue oAttrColl, attrLengthEndCone, dCone2Length

    
    'write back attributes for Can Rule Results, for reporting
    Set oAttrColl = GetAttributeCollection(oBUCan, InterfaceName_IJUASMCanRuleResult)

    SetAttributeValue oAttrColl, attrResultCanLength, dCanLength
    SetAttributeValue oAttrColl, attrResultUncutLength, dUncutLength
    SetAttributeValue oAttrColl, attrResultCanType, CanType_InLine
    SetAttributeValue oAttrColl, attrResultL2HullLength, dL2HullLength
    SetAttributeValue oAttrColl, attrResultL2CenterlineLength, dL2CenterlineLength
    SetAttributeValue oAttrColl, attrResultL3HullLength, dL3HullLength
    SetAttributeValue oAttrColl, attrResultL3CenterlineLength, dL3CenterlineLength
    SetAttributeValue oAttrColl, attrChamfer1Length, dChamferLengthStart
    SetAttributeValue oAttrColl, attrChamfer2Length, dChamferLengthEnd
    
    Dim sCanMaterial As String
    Dim sCanGrade As String
    Dim sCone1Material As String
    Dim sCone1Grade As String
    Dim sCone2Material As String
    Dim sCone2Grade As String
    
    With MyGC
        sCanMaterial = .Parameter("CanMaterial")
        sCanGrade = .Parameter("CanGrade")
        sCone1Material = .Parameter("Cone1Material")
        sCone1Grade = .Parameter("Cone1Grade")
        sCone2Material = .Parameter("Cone2Material")
        sCone2Grade = .Parameter("Cone2Grade")
    End With
    
    Set oAttrColl = GetAttributeCollection(oBUCan, InterfaceName_IUABuiltUpTube)
        
    SetAttributeValue oAttrColl, attrTubeThickness, dCanThickness
    SetAttributeValue oAttrColl, "TubeMaterial", sCanMaterial
    SetAttributeValue oAttrColl, "TubeGrade", sCanGrade

    Set oAttrColl = GetAttributeCollection(oBUCan, "IUABuiltUpCone1")
    SetAttributeValue oAttrColl, "Cone1Material", sCone1Material
    SetAttributeValue oAttrColl, "Cone1Grade", sCone1Grade
    SetAttributeValue oAttrColl, "Cone1Thickness", dCone1Thickness
    
    Set oAttrColl = GetAttributeCollection(oBUCan, "IUABuiltUpCone2")
    SetAttributeValue oAttrColl, "Cone2Material", sCone2Material
    SetAttributeValue oAttrColl, "Cone2Grade", sCone2Grade
    SetAttributeValue oAttrColl, "Cone2Thickness", dCone2Thickness
    
    Exit Sub
ErrorHandler:
    MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
    WriteToErrorLog E_FAIL, MODULE, METHOD, "Unspecified error"
    Err.Raise E_FAIL, MODULE & ":" & METHOD, "Unspecified error"
End Sub

' ****************************************************

Public Sub StubEndEvaluate(ByVal MyGC As SP3DGeometricConstruction.IJGeometricConstruction)
    Const METHOD = "StubEndEvaluate"
    On Error GoTo ErrorHandler
    
    ' use the GC as a GCMacro
    Dim iCanRule As ISPSCanRule
    Dim MyGCMacro As IJGeometricConstructionMacro
    Dim oBUCan As ISPSDesignedMember
    Dim crStatus As SPSCanRuleStatus

    Dim iPlane1 As IJPlane
    Dim dCanID As Double, dCanOD As Double
    Dim dNborStartID As Double, dNborStartOD As Double
    Dim dNborEndID As Double, dNborEndOD As Double
    Dim eTubeExtension As Long, eConeExtension As Long
    Dim dOverLengthStart As Double, dOverLengthEnd As Double
    Dim bIsHullStart As Boolean, bIsHullEnd As Boolean
    
    Dim dChamferLength As Double
    Dim dCone1Length As Double, dCone2Length As Double
    Dim dChamferSlope As Double, dFactor As Double
    Dim dMinimumExtensionDistance As Double
    Dim dRoundOffDistance As Double
    
    Dim posL2Hull As IJDPosition, posL2Centerline As IJDPosition
    Dim posL3Hull As IJDPosition, posL3Centerline As IJDPosition
    Dim oPrimaryMS As ISPSMemberSystem
    Dim vecPrimaryX As Double, vecPrimaryY As Double, vecPrimaryZ As Double
    Dim oAttrColl As Object
    Dim ePortId As SPSMemberAxisPortIndex

    Set iCanRule = MyGC
    Set MyGCMacro = MyGC
    
    ePortId = iCanRule.portIndex
    
    Set oPrimaryMS = MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Primary).Item("1")
    Set iPlane1 = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_Planes).Item("1") ' only one split plane for the end can
    Set oBUCan = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Item("1")
    Dim dCanThickness As Double

    ' dDiameterCan is already set on the BUCan by CrossSectionNotify semantic calling UpdateOutputCrossSectionDimensions.
    
    GetDiameters MyGCMacro, SPSMemberAxisAlong, dCanID, dCanOD
    'deosn't matter what port you pass below as there is only one neighbor for the end can
    GetDiameters MyGCMacro, SPSMemberAxisStart, dNborStartID, dNborStartOD
        
    eTubeExtension = MyGC.Parameter(attrL2Method)

    If eTubeExtension = TubeExtension_CLLength Then
        bIsHullStart = False
        dOverLengthStart = MyGC.Parameter(attrL2Length)
        dFactor = dOverLengthStart / dCanOD
        MyGC.Parameter(attrL2Factor) = dFactor

    ElseIf eTubeExtension = TubeExtension_CLFactor Then
        bIsHullStart = False
        dFactor = MyGC.Parameter(attrL2Factor)
        dOverLengthStart = dCanOD * dFactor
        MyGC.Parameter(attrL2Length) = dOverLengthStart

    ElseIf eTubeExtension = TubeExtension_HullLength Then
        bIsHullStart = True
        dOverLengthStart = MyGC.Parameter(attrL2Length)
        dFactor = dOverLengthStart / dCanOD
        MyGC.Parameter(attrL2Factor) = dFactor

    ElseIf eTubeExtension = TubeExtension_HullFactor Then
        bIsHullStart = True
        dFactor = MyGC.Parameter(attrL2Factor)
        dOverLengthStart = dCanOD * dFactor
        MyGC.Parameter(attrL2Length) = dOverLengthStart

    Else
        MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "error in tube extension parameter"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "error in tube extension parameter"
    End If

    dChamferSlope = MyGC.Parameter(attrChamferSlope)

'   Calculate Cone1Length and ChamferLength
    dCone1Length = 0#
    Dim dConeAngle As Double, dConeSlope As Double
    Dim dConeThickness As Double, dConeThicknessParameter As Double
    Dim dAdjustedConeThickness As Double

''    FindActualPlateThickness MyGC, "Can", dCanThickness       can thickness updated by ISPSCanRuleHelper_UpdateOutputCrossSectionDimensions
    dCanThickness = MyGC.Parameter(attrCanThickness)
    
    If ((dCanID >= dNborStartOD) Or (dCanOD <= dNborStartID)) Then 'Cone1 is required

        FindActualPlateThickness MyGC, "Cone", dConeThickness
        
        dConeThicknessParameter = MyGC.Parameter(attrConeThickness)
        If Abs(dConeThicknessParameter - dConeThickness) > dTol Then
            MyGC.Parameter(attrConeThickness) = dConeThickness
        End If
        
        eConeExtension = MyGC.Parameter(attrConeMethod)
    
        If eConeExtension = ConeMethod_Length Then
            dCone1Length = Abs(MyGC.Parameter(attrConeLength))
            
            If dCone1Length < dTol Then
                dConeAngle = 0
                dConeSlope = 0
                dCone1Length = 0
                dAdjustedConeThickness = dConeThickness
    
            Else
                ComputeConeAngle dCanOD, dCanThickness, dCone1Length, dNborStartOD, dConeThickness, dConeAngle, dAdjustedConeThickness
                dConeSlope = 1# / Tan(dConeAngle)
            
            End If

            MyGC.Parameter(attrConeSlope) = dConeSlope
            MyGC.Parameter(attrConeAngle) = dConeAngle
        
        ElseIf eConeExtension = ConeMethod_Slope Then
            dConeSlope = Abs(MyGC.Parameter(attrConeSlope))

            If dConeSlope > 100 Or dConeSlope < distTol Then
                dCone1Length = 0
                dConeAngle = 0
                dAdjustedConeThickness = dConeThickness

            Else
                dConeAngle = Atn(1# / dConeSlope)
                dAdjustedConeThickness = dConeThickness / Cos(dConeAngle)
                ComputeConeLength dConeAngle, dNborStartOD, dCanOD, dAdjustedConeThickness, dCanThickness, dCone1Length

            End If

            MyGC.Parameter(attrConeLength) = dCone1Length
            MyGC.Parameter(attrConeAngle) = dConeAngle

        ElseIf eConeExtension = ConeMethod_Angle Then
            dConeAngle = Abs(MyGC.Parameter(attrConeAngle))
            If dConeAngle < oneDegree Or dConeAngle > eightyNineDegrees Then
                dCone1Length = 0
                dConeSlope = 0
                dAdjustedConeThickness = dConeThickness

            Else
                dAdjustedConeThickness = dConeThickness / Cos(dConeAngle)
                dConeSlope = 1# / Tan(dConeAngle)
                ComputeConeLength dConeAngle, dNborStartOD, dCanOD, dAdjustedConeThickness, dCanThickness, dCone1Length

            End If

            MyGC.Parameter(attrConeLength) = dCone1Length
            MyGC.Parameter(attrConeSlope) = dConeSlope
        
        Else
            MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
            WriteToErrorLog E_FAIL, MODULE, METHOD, "error in cone extension parameter (Length)"
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "error in cone extension parameter (Length)"
        End If
        
        ComputeChamferLengthAtCone dChamferSlope, dCanThickness, dAdjustedConeThickness, dCone1Length, dChamferLength

    Else

        ComputeChamferLengthAtNbor dChamferSlope, dCanOD, dCanID, dNborStartOD, dNborStartID, dChamferLength

        dCone1Length = 0#
        dConeThickness = dCanThickness
    End If

'   Calculate Cone2Length and ChamferEndLength
    dCone2Length = 0# ' as there is only one cone for the StubEnd can

    GetUnitVectorTangent oPrimaryMS, vecPrimaryX, vecPrimaryY, vecPrimaryZ
    
    iCanRule.Services.ComputeMinMaxPoints 0, posL2Hull, posL2Centerline, posL3Centerline, posL3Hull, crStatus
    If crStatus <> SPSCanRule_Ok Then
        MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "problem computing hull length points"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "problem computing hull length points"
    End If

    Dim iIJPoint As IJPoint
    
    'set the position to the middle of the two points returned.
    Set iIJPoint = MyGCMacro
    iIJPoint.SetPoint 0.5 * (posL2Hull.x + posL3Hull.x), 0.5 * (posL2Hull.y + posL3Hull.y), 0.5 * (posL2Hull.z + posL3Hull.z)
    
    'Attributes for IJUASMCanRuleResult
    Dim dCanLength As Double
    Dim dUncutLength As Double
    Dim dL2HullLength As Double
    Dim dL2CenterlineLength As Double
    Dim dL3HullLength As Double     ' always zero for StubEnd
    Dim dL3CenterlineLength As Double   ' always zero for StubEnd
    
    Dim dL2Diff As Double
    Dim dL3Diff As Double
    Dim dLDiff As Double
    Dim dRoundOffDiff As Double
    
    dL2Diff = (posL2Centerline.x - posL2Hull.x) * vecPrimaryX + _
                (posL2Centerline.y - posL2Hull.y) * vecPrimaryY + _
                   (posL2Centerline.z - posL2Hull.z) * vecPrimaryZ
    dL3Diff = (posL3Hull.x - posL3Centerline.x) * vecPrimaryX + _
                (posL3Hull.y - posL3Centerline.y) * vecPrimaryY + _
                   (posL3Hull.z - posL3Centerline.z) * vecPrimaryZ
    
    'if the can is on the start of the member, use L3 values to decide split point
    'if it is on the end use L2 values to decide the split point
    If ePortId = SPSMemberAxisStart Then
        dLDiff = dL3Diff
    Else
        dLDiff = dL2Diff
    End If

    dMinimumExtensionDistance = MinimumExtensionDistance(CanType_StubEnd, dCanOD, MyGC)

    dUncutLength = 0# 'initialize
    
    If bIsHullStart Then
        dL2HullLength = dOverLengthStart
    Else
        dL2HullLength = dOverLengthStart - dLDiff
    End If
    If dL2HullLength < dMinimumExtensionDistance Then
        MyGC.PostError INVALID_PARAMETERS, False, strCodeListTablename, oBUCan
        ' the codelist simply says the parameter is invalid and refers the user to the error
        ' log, so writh the additional information needed
        WriteToErrorLog S_FALSE, MODULE, METHOD, "The specified " & attrL2Length & " length of " & Format$(dL2HullLength, "#.######") & _
                                                 " does not meet the specified " & _
                                                 " minimum of " & Format$(dMinimumExtensionDistance, "#.######")
        dL2HullLength = dMinimumExtensionDistance
    End If
   
    ' CanLength is based on the distance to the hull-hull "short" point, while
    ' UncutCanLength goes to the hull-hull "long" point.
    dUncutLength = posL2Hull.DistPt(posL3Hull) + dL2HullLength + dCone1Length + dChamferLength

    'round off uncut length
    dRoundOffDistance = MyGC.Parameter(attrRoundoffDistance)
    Dim dRoundOffUncutLength As Double
    If dRoundOffDistance > distTol Then
        dRoundOffUncutLength = Int((dUncutLength + dRoundOffDistance - dTol) / dRoundOffDistance) * dRoundOffDistance
        dRoundOffDiff = dRoundOffUncutLength - dUncutLength
        dUncutLength = dRoundOffUncutLength
    Else
        dRoundOffUncutLength = dUncutLength
        dRoundOffDiff = 0
    End If
    
    dCanLength = dUncutLength - dCone1Length
    
    'L3HullLength is always zero for StubEndCan, all round off diff should go to L2HullLength
    dL2HullLength = dL2HullLength + dRoundOffDiff
    dL2CenterlineLength = dL2HullLength + dLDiff
    
    If ePortId = SPSMemberAxisEnd Then
        posL2Hull.x = posL2Hull.x - (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryX
        posL2Hull.y = posL2Hull.y - (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryY
        posL2Hull.z = posL2Hull.z - (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryZ
        dNborEndOD = dCanOD

    Else
       
        posL2Hull.x = posL3Hull.x + (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryX
        posL2Hull.y = posL3Hull.y + (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryY
        posL2Hull.z = posL3Hull.z + (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryZ
                
        'only one code for the end cans. But need to swap the values as we need to set them correctly for the BUCan
        dCone2Length = dCone1Length
        dCone1Length = 0
        dNborEndOD = dNborStartOD
        dNborStartOD = dCanOD
    End If
    
    If Not PositionWithinRange(oPrimaryMS, posL2Hull) Then
        MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "hull length points off member system"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "hull length points off member system"
    End If
    
    UpdatePlane iPlane1, posL2Hull, vecPrimaryX, vecPrimaryY, vecPrimaryZ
    
    Set oAttrColl = GetAttributeCollection(oBUCan, InterfaceName_IUABuiltUpCan)

    SetAttributeValue oAttrColl, attrDiameterStart, dNborStartOD
    SetAttributeValue oAttrColl, attrLengthStartCone, dCone1Length
    SetAttributeValue oAttrColl, attrDiameterEnd, dNborEndOD
    SetAttributeValue oAttrColl, attrLengthEndCone, dCone2Length

    Set oAttrColl = GetAttributeCollection(oBUCan, InterfaceName_IJUASMCanRuleResult)

    SetAttributeValue oAttrColl, attrResultCanLength, dCanLength
    SetAttributeValue oAttrColl, attrResultUncutLength, dUncutLength
    SetAttributeValue oAttrColl, attrResultCanType, CanType_StubEnd
    SetAttributeValue oAttrColl, attrResultL2HullLength, dL2HullLength
    SetAttributeValue oAttrColl, attrResultL2CenterlineLength, dL2CenterlineLength
    SetAttributeValue oAttrColl, attrResultL3HullLength, dL3HullLength
    SetAttributeValue oAttrColl, attrResultL3CenterlineLength, dL3CenterlineLength
    If ePortId = SPSMemberAxisEnd Then
        SetAttributeValue oAttrColl, attrChamfer1Length, dChamferLength
        SetAttributeValue oAttrColl, attrChamfer2Length, 0#
    Else
        SetAttributeValue oAttrColl, attrChamfer1Length, 0#
        SetAttributeValue oAttrColl, attrChamfer2Length, dChamferLength
    End If
    
    
    Dim sConeMaterial As String
    Dim sConeGrade As String
    Dim sCanMaterial As String
    Dim sCanGrade As String
    
    With MyGC
        sCanMaterial = .Parameter("CanMaterial")
        sCanGrade = .Parameter("CanGrade")
        sConeMaterial = .Parameter("ConeMaterial")
        sConeGrade = .Parameter("ConeGrade")
    End With
    
     Set oAttrColl = GetAttributeCollection(oBUCan, InterfaceName_IUABuiltUpTube)
        
     SetAttributeValue oAttrColl, attrTubeThickness, dCanThickness
     SetAttributeValue oAttrColl, "TubeMaterial", sCanMaterial
     SetAttributeValue oAttrColl, "TubeGrade", sCanGrade
     
     Dim dCone1Thickness As Double
     Dim sCone1Material As String
     Dim sCone1Grade As String
     Dim dCone2Thickness As Double
     Dim sCone2Material As String
     Dim sCone2Grade As String
     
     If ePortId = SPSMemberAxisEnd Then
         dCone1Thickness = dConeThickness
         sCone1Material = sConeMaterial
         sCone1Grade = sConeGrade
         dCone2Thickness = dCanThickness
         sCone2Material = sCanMaterial
         sCone2Grade = sCanGrade
    Else
         dCone1Thickness = dCanThickness
         sCone1Material = sCanMaterial
         sCone1Grade = sCanGrade
         dCone2Thickness = dConeThickness
         sCone2Material = sConeMaterial
         sCone2Grade = sConeGrade
    End If
         
    Set oAttrColl = GetAttributeCollection(oBUCan, "IUABuiltUpCone1")
    SetAttributeValue oAttrColl, "Cone1Thickness", dCone1Thickness
    SetAttributeValue oAttrColl, "Cone1Material", sCone1Material
    SetAttributeValue oAttrColl, "Cone1Grade", sCone1Grade
    Set oAttrColl = GetAttributeCollection(oBUCan, "IUABuiltUpCone2")
    SetAttributeValue oAttrColl, "Cone2Thickness", dCone2Thickness
    SetAttributeValue oAttrColl, "Cone2Material", sCone2Material
    SetAttributeValue oAttrColl, "Cone2Grade", sCone2Grade

    Exit Sub

ErrorHandler:
    MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
    WriteToErrorLog E_FAIL, MODULE, METHOD, "Unspecified error"
    Err.Raise E_FAIL, MODULE & ":" & METHOD, "Unspecified error"
End Sub

' *********************************************************

Public Sub EndEvaluate(ByVal MyGC As SP3DGeometricConstruction.IJGeometricConstruction)
    Const METHOD = "EndEvaluate"
    On Error GoTo ErrorHandler
    
    ' use the GC as a GCMacro
    Dim iCanRule As ISPSCanRule
    Dim MyGCMacro As IJGeometricConstructionMacro
    Dim oBUCan As ISPSDesignedMember
    Dim crStatus As SPSCanRuleStatus

    Dim iPlane1 As IJPlane
    Dim iOffsetPlane As IJPlane
    Dim dCanID As Double, dCanOD As Double
    Dim dNborStartID As Double, dNborStartOD As Double
    Dim dNborEndID As Double, dNborEndOD As Double
    Dim eTubeMethodL2 As Long, eTubeMethodL3 As Long, eConeMethod As Long
    Dim dOverLengthStart As Double, dOverLengthEnd As Double
    Dim bIsHullStart As Boolean, bIsHullEnd As Boolean
    
    Dim dChamferLength As Double
    Dim dCone1Length As Double, dCone2Length As Double
    Dim dChamferSlope As Double, dFactor As Double
    Dim dMinimumExtensionDistance As Double
    Dim dRoundOffDistance As Double
    
    Dim posL2Hull As IJDPosition, posL2Centerline As IJDPosition
    Dim posL3Hull As IJDPosition, posL3Centerline As IJDPosition
    Dim oPrimaryMS As ISPSMemberSystem
    Dim vecPrimaryX As Double, vecPrimaryY As Double, vecPrimaryZ As Double
    Dim oAttrColl As Object
    Dim ePortId As SPSMemberAxisPortIndex
    Dim dCanThickness As Double
    Dim dAdjustedConeThickness As Double

    Set iCanRule = MyGC
    Set MyGCMacro = MyGC
    
    ePortId = iCanRule.portIndex
    
    Set oPrimaryMS = MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Primary).Item("1")
    Set iPlane1 = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_Planes).Item("1") ' only one split plane for the end can
    Set iOffsetPlane = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_OffsetSurface).Item("1")
    Set oBUCan = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Item("1")
    
    ' dDiameterCan is already set on the BUCan by CrossSectionNotify semantic calling UpdateOutputCrossSectionDimensions.
    ' treat endCan as though we are at the endPort, which means the cone is the startCone of the BUCan
    ' later, we'll set the right stuff.

    GetDiameters MyGCMacro, SPSMemberAxisAlong, dCanID, dCanOD
    GetDiameters MyGCMacro, SPSMemberAxisStart, dNborStartID, dNborStartOD

    eTubeMethodL2 = MyGC.Parameter(attrL2Method)

    If eTubeMethodL2 = TubeExtension_CLLength Then
        bIsHullStart = False
        dOverLengthStart = MyGC.Parameter(attrL2Length)
        dFactor = dOverLengthStart / dCanOD
        MyGC.Parameter(attrL2Factor) = dFactor

    ElseIf eTubeMethodL2 = TubeExtension_CLFactor Then
        bIsHullStart = False
        dFactor = MyGC.Parameter(attrL2Factor)
        dOverLengthStart = dCanOD * dFactor
        MyGC.Parameter(attrL2Length) = dOverLengthStart
    
    ElseIf eTubeMethodL2 = TubeExtension_HullLength Then
        bIsHullStart = True
        dOverLengthStart = MyGC.Parameter(attrL2Length)
        dFactor = dOverLengthStart / dCanOD
        MyGC.Parameter(attrL2Factor) = dFactor

    ElseIf eTubeMethodL2 = TubeExtension_HullFactor Then
        bIsHullStart = True
        dFactor = MyGC.Parameter(attrL2Factor)
        dOverLengthStart = dCanOD * dFactor
        MyGC.Parameter(attrL2Length) = dOverLengthStart

    Else
        MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "error in tube extension parameter"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "error in tube extension parameter"
    End If
    
    eTubeMethodL3 = MyGC.Parameter(attrL3Method)

    If eTubeMethodL3 = TubeExtension_CLLength Then
        bIsHullEnd = False
        dOverLengthEnd = MyGC.Parameter(attrL3Length)
        dFactor = dOverLengthEnd / dCanOD
        MyGC.Parameter(attrL3Factor) = dFactor

    ElseIf eTubeMethodL3 = TubeExtension_CLFactor Then
        bIsHullEnd = False
        dFactor = MyGC.Parameter(attrL3Factor)
        dOverLengthEnd = dCanOD * dFactor
        MyGC.Parameter(attrL3Length) = dOverLengthEnd

    ElseIf eTubeMethodL3 = TubeExtension_HullLength Then
        bIsHullEnd = True
        dOverLengthEnd = MyGC.Parameter(attrL3Length)
        dFactor = dOverLengthEnd / dCanOD
        MyGC.Parameter(attrL3Factor) = dFactor

    ElseIf eTubeMethodL3 = TubeExtension_HullFactor Then
        bIsHullEnd = True
        dFactor = MyGC.Parameter(attrL3Factor)
        dOverLengthEnd = dCanOD * dFactor
        MyGC.Parameter(attrL3Length) = dOverLengthEnd

    Else
        MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "error in tube extension parameter"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "error in tube extension parameter"
    End If

    dChamferSlope = MyGC.Parameter(attrChamferSlope)

'   Calculate Cone1Length and ChamferStartLength
    dCone1Length = 0#
    Dim dConeAngle As Double, dConeSlope As Double
    Dim dConeThickness As Double, dConeThicknessParameter As Double
    Dim diffRadius As Double

''    FindActualPlateThickness MyGC, "Can", dCanThickness       can thickness updated by ISPSCanRuleHelper_UpdateOutputCrossSectionDimensions
    dCanThickness = MyGC.Parameter(attrCanThickness)

    If ((dCanID >= dNborStartOD) Or (dCanOD <= dNborStartID)) Then 'Cone1 is required

        FindActualPlateThickness MyGC, "Cone", dConeThickness
        
        dConeThicknessParameter = MyGC.Parameter(attrConeThickness)
        If Abs(dConeThicknessParameter - dConeThickness) > dTol Then
            MyGC.Parameter(attrConeThickness) = dConeThickness
        End If
    
        eConeMethod = MyGC.Parameter(attrConeMethod)
    
        If eConeMethod = ConeMethod_Length Then
            dCone1Length = Abs(MyGC.Parameter(attrConeLength))
            
            If dCone1Length < dTol Then
                dConeAngle = 0
                dConeSlope = 0
                dCone1Length = 0
                dAdjustedConeThickness = dConeThickness
    
            Else
                ComputeConeAngle dCanOD, dCanThickness, dCone1Length, dNborStartOD, dConeThickness, dConeAngle, dAdjustedConeThickness
                dConeSlope = 1# / Tan(dConeAngle)
            
            End If

            MyGC.Parameter(attrConeSlope) = dConeSlope
            MyGC.Parameter(attrConeAngle) = dConeAngle
        
        ElseIf eConeMethod = ConeMethod_Slope Then
            dConeSlope = Abs(MyGC.Parameter(attrConeSlope))

            If dConeSlope > 100 Or dConeSlope < distTol Then
                dCone1Length = 0
                dConeAngle = 0
                dAdjustedConeThickness = dConeThickness

            Else
                dConeAngle = Atn(1# / dConeSlope)
                dAdjustedConeThickness = dConeThickness / Cos(dConeAngle)
                ComputeConeLength dConeAngle, dNborStartOD, dCanOD, dAdjustedConeThickness, dCanThickness, dCone1Length

            End If

            MyGC.Parameter(attrConeLength) = dCone1Length
            MyGC.Parameter(attrConeAngle) = dConeAngle
        
        ElseIf eConeMethod = ConeMethod_Angle Then
            dConeAngle = Abs(MyGC.Parameter(attrConeAngle))
            If dConeAngle < oneDegree Or dConeAngle > eightyNineDegrees Then
                dCone1Length = 0
                dConeSlope = 0
                dAdjustedConeThickness = dConeThickness

            Else
                dAdjustedConeThickness = dConeThickness / Cos(dConeAngle)
                dConeSlope = 1# / Tan(dConeAngle)
                ComputeConeLength dConeAngle, dNborStartOD, dCanOD, dAdjustedConeThickness, dCanThickness, dCone1Length

            End If

            MyGC.Parameter(attrConeLength) = dCone1Length
            MyGC.Parameter(attrConeSlope) = dConeSlope
        
        Else
            MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
            WriteToErrorLog E_FAIL, MODULE, METHOD, "error in cone extension parameter (Length)"
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "error in cone extension parameter (Length)"
        End If
        
        ComputeChamferLengthAtCone dChamferSlope, dCanThickness, dAdjustedConeThickness, dCone1Length, dChamferLength

    Else
    
        ComputeChamferLengthAtNbor dChamferSlope, dCanOD, dCanID, dNborStartOD, dNborStartID, dChamferLength
        
        dCone1Length = 0#
        dConeThickness = dCanThickness
    End If

    dCone2Length = 0# ' as there is only one cone for the endcan

    GetUnitVectorTangent oPrimaryMS, vecPrimaryX, vecPrimaryY, vecPrimaryZ
    
    iCanRule.Services.ComputeMinMaxPoints 1, posL2Hull, posL2Centerline, posL3Centerline, posL3Hull, crStatus
    If crStatus <> SPSCanRule_Ok Then
        MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "problem computing hull length points"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "problem computing hull length points"
    End If

    Dim iIJPoint As IJPoint
    
    'set the position to the middle of the two points returned.
    Set iIJPoint = MyGCMacro
    iIJPoint.SetPoint 0.5 * (posL2Hull.x + posL3Hull.x), 0.5 * (posL2Hull.y + posL3Hull.y), 0.5 * (posL2Hull.z + posL3Hull.z)
    
    'Attributes for IJUASMCanRuleResult
    Dim dCanLength As Double
    Dim dUncutLength As Double
    Dim dL2HullLength As Double
    Dim dL2CenterlineLength As Double
    Dim dL3HullLength As Double
    Dim dL3CenterlineLength As Double
    
    Dim dL2Diff As Double
    Dim dL3Diff As Double
    Dim dRoundOffDiff As Double
    
    dL2Diff = (posL2Centerline.x - posL2Hull.x) * vecPrimaryX + _
                 (posL2Centerline.y - posL2Hull.y) * vecPrimaryY + _
                    (posL2Centerline.z - posL2Hull.z) * vecPrimaryZ
    dL3Diff = (posL3Hull.x - posL3Centerline.x) * vecPrimaryX + _
                 (posL3Hull.y - posL3Centerline.y) * vecPrimaryY + _
                    (posL3Hull.z - posL3Centerline.z) * vecPrimaryZ
    
    dMinimumExtensionDistance = MinimumExtensionDistance(CanType_End, dCanOD, MyGC)

    dUncutLength = 0# 'initialize
    
    If bIsHullStart Then
        dL2HullLength = dOverLengthStart
    Else
        dL2HullLength = dOverLengthStart - dL2Diff
    End If
    If dL2HullLength < dMinimumExtensionDistance Then
        MyGC.PostError INVALID_PARAMETERS, False, strCodeListTablename, oBUCan
        ' the codelist simply says the parameter is invalid and refers the user to the error
        ' log, so writh the additional information needed
        WriteToErrorLog S_FALSE, MODULE, METHOD, "The specified " & attrL2Length & " length of " & Format$(dL2HullLength, "#.######") & _
                                                 " does not meet the specified " & _
                                                 " minimum of " & Format$(dMinimumExtensionDistance, "#.######")

        dL2HullLength = dMinimumExtensionDistance
    End If
   
    If bIsHullEnd Then
        dL3HullLength = dOverLengthEnd
    Else
        dL3HullLength = dOverLengthEnd - dL3Diff
    End If
    If dL3HullLength < dMinimumExtensionDistance Then
        MyGC.PostError INVALID_PARAMETERS, False, strCodeListTablename, oBUCan
        ' the codelist simply says the parameter is invalid and refers the user to the error
        ' log, so writh the additional information needed
        WriteToErrorLog S_FALSE, MODULE, METHOD, "The specified " & attrL3Length & " length of " & Format$(dL3HullLength, "#.######") & _
                                                 " does not meet the specified " & _
                                                 " minimum of " & Format$(dMinimumExtensionDistance, "#.######")
        dL3HullLength = dMinimumExtensionDistance
    End If
    
    dUncutLength = Abs((posL3Hull.x - posL2Hull.x) * vecPrimaryX + (posL3Hull.y - posL2Hull.y) * vecPrimaryY + (posL3Hull.z - posL2Hull.z) * vecPrimaryZ)
    dUncutLength = dUncutLength + dL2HullLength + dL3HullLength + dCone1Length + dChamferLength
    
    'round off uncut length
    dRoundOffDistance = MyGC.Parameter(attrRoundoffDistance)
    Dim dRoundOffUncutLength As Double
    If dRoundOffDistance > distTol Then
        dRoundOffUncutLength = Int((dUncutLength + dRoundOffDistance - dTol) / dRoundOffDistance) * dRoundOffDistance
        dRoundOffDiff = dRoundOffUncutLength - dUncutLength
        dUncutLength = dRoundOffUncutLength
    Else
        dRoundOffUncutLength = dUncutLength
        dRoundOffDiff = 0
    End If

    dL2HullLength = dL2HullLength + dRoundOffDiff / 2
    dL2CenterlineLength = dL2HullLength + dL2Diff
    dL3HullLength = dL3HullLength + dRoundOffDiff / 2
    dL3CenterlineLength = dL3HullLength + dL3Diff
    

    Dim posSplitLocation As IJDPosition
    Set posSplitLocation = New DPosition
    Dim posOffsetLocation As IJDPosition
    Set posOffsetLocation = New DPosition
    
    ' L2 is applied closest to the cone, L3 is an extension applied on the outside, beyond secondary, if present.
    ' posL2Hull is minimum along the axis, and posL3Hull is maximum along the axis.
    
    If ePortId = SPSMemberAxisEnd Then
        ' move split location inboard from L2Hull
        posSplitLocation.x = posL2Hull.x - (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryX
        posSplitLocation.y = posL2Hull.y - (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryY
        posSplitLocation.z = posL2Hull.z - (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryZ
        
        ' move offset outboard of the L3 position
        posOffsetLocation.x = posL3Hull.x + (dL3HullLength) * vecPrimaryX
        posOffsetLocation.y = posL3Hull.y + (dL3HullLength) * vecPrimaryY
        posOffsetLocation.z = posL3Hull.z + (dL3HullLength) * vecPrimaryZ

    Else
        ' move split location inboard from the L3Hull pos, using L2 length
        posSplitLocation.x = posL3Hull.x + (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryX
        posSplitLocation.y = posL3Hull.y + (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryY
        posSplitLocation.z = posL3Hull.z + (dL2HullLength + dCone1Length + dChamferLength) * vecPrimaryZ
        
        ' move offset outboard L2Hull position by L3 distance.
        posOffsetLocation.x = posL2Hull.x - (dL3HullLength) * vecPrimaryX
        posOffsetLocation.y = posL2Hull.y - (dL3HullLength) * vecPrimaryY
        posOffsetLocation.z = posL2Hull.z - (dL3HullLength) * vecPrimaryZ

        'only one code for the end cans. But need to swap the values as we need to set them correctly for the BUCAN
        dCone2Length = dCone1Length
        dCone1Length = 0
        dNborEndOD = dNborStartOD
        
    End If

    dCanLength = dUncutLength - IIf(ePortId = SPSMemberAxisEnd, dCone1Length, dCone2Length)

    If Not PositionWithinRange(oPrimaryMS, posSplitLocation) Then
        MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "hull length points off member system"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "hull length points off member system"
    End If

    UpdatePlane iPlane1, posSplitLocation, vecPrimaryX, vecPrimaryY, vecPrimaryZ
    
    UpdatePlane iOffsetPlane, posOffsetLocation, vecPrimaryX, vecPrimaryY, vecPrimaryZ

    Set oAttrColl = GetAttributeCollection(oBUCan, InterfaceName_IUABuiltUpCan)
    
    SetAttributeValue oAttrColl, attrDiameterStart, IIf(dCone1Length < 0.001, dCanOD, dNborStartOD)
    SetAttributeValue oAttrColl, attrLengthStartCone, dCone1Length
    
    SetAttributeValue oAttrColl, attrDiameterEnd, IIf(dCone2Length < 0.001, dCanOD, dNborStartOD)
    SetAttributeValue oAttrColl, attrLengthEndCone, dCone2Length

    Set oAttrColl = GetAttributeCollection(oBUCan, InterfaceName_IJUASMCanRuleResult)

    SetAttributeValue oAttrColl, attrResultCanLength, dCanLength
    SetAttributeValue oAttrColl, attrResultUncutLength, dUncutLength
    SetAttributeValue oAttrColl, attrResultCanType, CanType_End
    SetAttributeValue oAttrColl, attrResultL2HullLength, dL2HullLength
    SetAttributeValue oAttrColl, attrResultL2CenterlineLength, dL2CenterlineLength
    SetAttributeValue oAttrColl, attrResultL3HullLength, dL3HullLength
    SetAttributeValue oAttrColl, attrResultL3CenterlineLength, dL3CenterlineLength
    
    If ePortId = SPSMemberAxisEnd Then
        SetAttributeValue oAttrColl, attrChamfer1Length, dChamferLength
        SetAttributeValue oAttrColl, attrChamfer2Length, 0#
    Else
        SetAttributeValue oAttrColl, attrChamfer1Length, 0#
        SetAttributeValue oAttrColl, attrChamfer2Length, dChamferLength
    End If
    
    Dim sConeMaterial As String
    Dim sConeGrade As String
    Dim sCanMaterial As String
    Dim sCanGrade As String
    
    With MyGC
        sCanMaterial = .Parameter("CanMaterial")
        sCanGrade = .Parameter("CanGrade")
        sConeMaterial = .Parameter("ConeMaterial")
        sConeGrade = .Parameter("ConeGrade")
    End With
    
     Set oAttrColl = GetAttributeCollection(oBUCan, InterfaceName_IUABuiltUpTube)
        
     SetAttributeValue oAttrColl, attrTubeThickness, dCanThickness
     SetAttributeValue oAttrColl, "TubeMaterial", sCanMaterial
     SetAttributeValue oAttrColl, "TubeGrade", sCanGrade
     
     Dim dCone1Thickness As Double
     Dim sCone1Material As String
     Dim sCone1Grade As String
     Dim dCone2Thickness As Double
     Dim sCone2Material As String
     Dim sCone2Grade As String
     
     If ePortId = SPSMemberAxisEnd Then         ' the BUCan will see this as a startCone
         dCone1Thickness = dConeThickness
         sCone1Material = sConeMaterial
         sCone1Grade = sConeGrade
         dCone2Thickness = dCanThickness
         sCone2Material = sCanMaterial
         sCone2Grade = sCanGrade
    Else                                        ' the BUCan will see this as an endCone
         dCone1Thickness = dCanThickness
         sCone1Material = sCanMaterial
         sCone1Grade = sCanGrade
         dCone2Thickness = dConeThickness
         sCone2Material = sConeMaterial
         sCone2Grade = sConeGrade
    End If
     
    Set oAttrColl = GetAttributeCollection(oBUCan, "IUABuiltUpCone1")
    SetAttributeValue oAttrColl, "Cone1Thickness", dCone1Thickness
    SetAttributeValue oAttrColl, "Cone1Material", sCone1Material
    SetAttributeValue oAttrColl, "Cone1Grade", sCone1Grade

    Set oAttrColl = GetAttributeCollection(oBUCan, "IUABuiltUpCone2")
    SetAttributeValue oAttrColl, "Cone2Thickness", dCone2Thickness
    SetAttributeValue oAttrColl, "Cone2Material", sCone2Material
    SetAttributeValue oAttrColl, "Cone2Grade", sCone2Grade
    
    Exit Sub

ErrorHandler:
    MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
    WriteToErrorLog E_FAIL, MODULE, METHOD, "Unspecified error"
    Err.Raise E_FAIL, MODULE & ":" & METHOD, "Unspecified error"
End Sub

'DI-166343 Minimum Extension Distance requirement enforced in catalog for well-defined can, also need to enforce the rule for customized can
Public Function MinimumExtensionDistance(lCanType As Long, CanOD As Double, MyGC As SP3DGeometricConstruction.IJGeometricConstruction) As Double
    Const METHOD = "MinimumExtensionDistance"
    On Error GoTo ErrorHandler

    Dim dMinExtensionBasedOnDiameter As Double
    Dim dMinExtensionParameter As Double
    
    dMinExtensionParameter = MyGC.Parameter(attrMinExtensionDistance)

    ' establish the minimum required extension based on Can type.  API standard.
    ' for InLine, use OD / 4.  For Stub and End, use D.

    If lCanType = CanType_InLine Then
        dMinExtensionBasedOnDiameter = CanOD / 4#
    Else
        dMinExtensionBasedOnDiameter = CanOD
    End If
    
    ' now use the larger of extension based on diameter or parameter value
    
    If dMinExtensionBasedOnDiameter > dMinExtensionParameter Then
        ' Post a warning that the specified extension
        ' either from the catalog or from user input does not
        ' meet the API minimum.
        Dim MyGCMacro As IJGeometricConstructionMacro
        Set MyGCMacro = MyGC
        MyGC.PostError INVALID_PARAMETERS, False, strCodeListTablename, MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Item("1")
        ' the codelist simply says the parameter is invalid and refers the user to the error
        ' log, so writh the additional information needed
        WriteToErrorLog S_FALSE, MODULE, METHOD, "The specified " & attrMinExtensionDistance & " length of " & Format$(dMinExtensionParameter, "#.######") & _
                                                 " does not meet the American Petroleum Institute's" & _
                                                 " minimum of " & Format$(dMinExtensionBasedOnDiameter, "#.######")
        Err.Clear
        
        MinimumExtensionDistance = dMinExtensionBasedOnDiameter
    Else
        MinimumExtensionDistance = dMinExtensionParameter
    End If

    Exit Function

ErrorHandler:
    WriteToErrorLog E_FAIL, MODULE, METHOD, "Unspecified error"
    Err.Raise E_FAIL, MODULE & ":" & METHOD, "Unspecified error"
End Function

'TR-166637 Calculate Cone angle from Cone Length and ODs on two cone ends
Public Function ConeAngleFromLengthAndOD(dLength As Double, dOD1 As Double, dOD2 As Double) As Double
    Const METHOD = "ConeAngleFromLengthAndOD"
    On Error GoTo ErrorHandler
    ConeAngleFromLengthAndOD = 0#
    If Abs(dLength) > dTol Then
        ConeAngleFromLengthAndOD = Abs(Atn((dOD1 - dOD2) / dLength / 2))
    End If
    Exit Function
ErrorHandler:
    WriteToErrorLog E_FAIL, MODULE, METHOD, "Unspecified error"
    Err.Raise E_FAIL, MODULE & ":" & METHOD, "Unspecified error"
End Function

' Nominally, the coneLength is simply a multiple of difference in diameter times slope; ie: (dCanOD - dNborStartOD)) * dConeSlope
' However, the BUCan computes the curve to revolve as the ID curve.  That is done by using the OD and subtracting adjusted cone-thickness and tube-thickness.
' This function computes the coneLength so that that ID curve will produce the desired cone slope.

Public Sub ComputeConeLength(dConeAngle As Double, dNborOD As Double, _
                dCanOD As Double, dAdjustedConeThickness As Double, dCanThickness As Double, ByRef dConeLength As Double)

    Dim dInnerConeLength As Double
    Dim diffInnerRadius As Double
    Dim dCanInnerRadius As Double, dConeInnerRadius As Double
    Dim dTanConeAngle As Double

    dTanConeAngle = Tan(Abs(dConeAngle))

    dCanInnerRadius = (0.5 * dCanOD) - dCanThickness
    dConeInnerRadius = (0.5 * dNborOD) - dAdjustedConeThickness

    diffInnerRadius = dCanInnerRadius - dConeInnerRadius

    dInnerConeLength = Abs(diffInnerRadius) / dTanConeAngle

    dConeLength = dInnerConeLength
    
End Sub
            

' compute cone angle and adjusted cone thickness based on that angle.
' adjusted cone thickness is the cone thickness / Cos ( coneAngle )
'
' But, the coneAngle must be the angle based on the ID's, so we iterate to find both the angle and the adjustedConeThickness

Public Sub ComputeConeAngle(dCanOD As Double, dCanThickness As Double, dConeLength As Double, dNborOD As Double, dConeThickness As Double, _
                                ByRef dConeAngle As Double, ByRef dAdjustedConeThickness As Double)

    Dim diffInnerRadius As Double
    
    dAdjustedConeThickness = dConeThickness             ' a good first approximation
    
    If dConeLength > dTol Then

        diffInnerRadius = 0.5 * Abs((dCanOD - 2 * dCanThickness) - (dNborOD - 2 * dAdjustedConeThickness))
        dConeAngle = Atn(diffInnerRadius / dConeLength)
        dAdjustedConeThickness = dConeThickness / Cos(dConeAngle)
    
        ' now using the adjustedConeThickness, recalculate the angle, and improve the adjustedConeThickness
        diffInnerRadius = 0.5 * Abs((dCanOD - 2 * dCanThickness) - (dNborOD - 2 * dAdjustedConeThickness))
        dConeAngle = Atn(diffInnerRadius / dConeLength)
        dAdjustedConeThickness = dConeThickness / Cos(dConeAngle)
    
        ' now using the adjustedConeThickness, recalculate the angle, and improve the adjustedConeThickness
        ' this converges extremely well.
        diffInnerRadius = 0.5 * Abs((dCanOD - 2 * dCanThickness) - (dNborOD - 2 * dAdjustedConeThickness))
        dConeAngle = Atn(diffInnerRadius / dConeLength)
        dAdjustedConeThickness = dConeThickness / Cos(dConeAngle)
    
    Else
    
        dConeAngle = 0
    
    End If

End Sub

Public Sub ComputeChamferLengthAtCone(dChamferSlope As Double, dCanThickness As Double, dAdjustedConeThickness As Double, _
                            dConeLength As Double, ByRef dChamferLength As Double)

    dChamferLength = 0

    If dConeLength > dTol Then
        If dChamferSlope > dTol And dChamferSlope <= 100 Then
            If dCanThickness > dAdjustedConeThickness Then
                dChamferLength = (dCanThickness - dAdjustedConeThickness) * dChamferSlope
            End If
        End If
    End If

End Sub

Public Sub ComputeChamferLengthAtNbor(dChamferSlope As Double, dCanOD As Double, dCanID As Double, _
                                        dNborOD As Double, dNborID As Double, ByRef dChamferLength As Double)
    
    Dim diffOD As Double, diffID As Double

    If dChamferSlope > dTol And dChamferSlope < 100 + dTol Then       ' similar test as for cone.

        diffOD = dCanOD - dNborOD       ' chamfer only if can OD is bigger
        diffID = dNborID - dCanID       ' chamfer only if nborID is bigger
        
        If diffOD > dTol And diffID > dTol Then     ' OD of can is larger, and ID of can is smaller than Nbor.  Use biggest difference.
            dChamferLength = 0.5 * dChamferSlope * IIf(diffOD > diffID, diffOD, diffID)

        ElseIf diffOD > dTol Then                   ' OD of can is larger than Nbor, and ID of can is bigger or same as nbor.
            dChamferLength = 0.5 * dChamferSlope * diffOD

        ElseIf diffID > dTol Then                   ' ID of Nbor is larger than can, and OD of can is smaller or same as nbor.
            dChamferLength = 0.5 * dChamferSlope * diffID
        
        Else
            dChamferLength = 0                      ' can is thinner or same as nbor.  no chamfer on can.

        End If
    
    Else
        dChamferLength = 0                          ' chamferSlope out of range.  no chamfer on can.
    
    End If
                            
End Sub

' function checks whether the given position is within valid parameter range of the given member system.
' When the curve is a line, "param" values are expressed in units of length.  This function includes a hard-coded
' minimum distance of 1 cm for closeness to the end of the memberSystem.

Private Function PositionWithinRange(iCurve As IJCurve, position As IJDPosition) As Boolean

    Dim dParPos As Double, dParStart As Double, dParEnd As Double
    Dim dDistMinDistToEnd As Double
    
    PositionWithinRange = False
    dDistMinDistToEnd = 0.01    ' one cm
    
    iCurve.ParamRange dParStart, dParEnd
    iCurve.Parameter position.x, position.y, position.z, dParPos
    
    If dParPos >= (dParStart + dDistMinDistToEnd) And dParPos <= (dParEnd - dDistMinDistToEnd) Then
        PositionWithinRange = True
    End If

End Function
