VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation. All rights reserved.
'
'  Project     : SharedContent\Src\Planning\Rules\CommonEntity\ComparisonRule\CommonEntityComparison.vbp
'  File        : ProfileRule.cls
'
'  Description : Common Entity Profile comparison rule
'
'  History     :
'   9th Jan 2012      Devi Kishore Adiraju     Initial creation
'**************************************************************************************


Option Explicit

Private Const IID_IJPlnProductionRouting As String = "{E9B1D9F6-C687-40BA-85E2-311534F4782A}"

Implements IJCommonEntityComparison
Implements IJStandardEntityComparison

Private Sub Class_Initialize()

    sSOURCEFILE = "ProfileRule.cls"

    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject(ERRORPROGID)
    End If

    m_dDistanceTolerance = 0.001
    m_dAreaTolerance = 0.001
    m_dVolumeTolerance = 0.001
    m_dWeightTolerance = 0.01
    m_dAngleTolerance = 0.001

End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

Private Function IJCommonEntityComparison_AreCommon(ByVal pCandidate As Object, ByVal pTarget As Object, strDisplayChecksColl() As String, strKeywordChecksColl() As String, ByVal pManager As Object, strResultMessage As String, saCandCheckValues() As Variant, saTgtCheckValues() As Variant, saStatus() As Variant) As Boolean
Const METHOD = "IJCommonEntityComparison_AreCommon"
On Error GoTo ErrorHandler

    Dim j                           As Long
    Dim oCandidate                  As IJProfilePart
    Dim oTarget                     As IJProfilePart
    Dim bCommon                     As Boolean

    Dim oCndtMaterial               As IJDMaterial
    Dim oTgtMaterial                As IJDMaterial
    Dim oCndtXSection               As IJCrossSection
    Dim oTgtXSection                As IJCrossSection
    Dim eCndtProfileType            As StructProfileType
    Dim eTgtProfileType             As StructProfileType
    Dim eCndtCurvature              As ProfileCurvature
    Dim eTgtCurvature               As ProfileCurvature
    Dim eCndtPrimOrientation        As StructMoldedOrientation
    Dim eTgtPrimOrientation         As StructMoldedOrientation
    Dim oCndtOrientVector           As IJDVector
    Dim oTgtOrientVector            As IJDVector
    Dim oConnectedObjPairs          As IJElements
    Dim oCndtConnectedObjects       As IJElements
    Dim oTgtConnectedObjects        As IJElements
    Dim oCndtProdRouting            As Object
    Dim oTgtProdRouting             As Object
    Dim oCndtProdRtgActions         As IJElements
    Dim oTgtProdRtgActions          As IJElements
    Dim oMatchedPCs                 As IJElements
    Dim bPCsAlreadyMatched          As Boolean

    Dim dCndt_X As Double, dCndt_Y As Double, dCndt_Z As Double
    Dim dTgt_X As Double, dTgt_Y As Double, dTgt_Z As Double

    Dim dCndtLndCrvLength           As Double
    Dim dTgtLndCrvLength            As Double

    Dim vCndtPropValue              As Variant
    Dim vTgtPropValue               As Variant
    Dim bAPISuccess                 As Boolean
    Dim bIsGeomSame                 As Boolean
    Dim strBevelInfo                As String

    'Create a Common Part Plate Helper and Common Helper

    Dim oCPSCommonHelper            As IJPlnCompareHelperEx
    Dim oCPSProfileHelper             As IJPlnProfileHelper

    Set oCPSProfileHelper = New CPlnProfileHelper
    Set oCPSCommonHelper = oCPSProfileHelper

    'By default the comparison returns true
    bCommon = True
    m_bGeometriesCompared = False   'Indicates that the geometries are not compared yet.
    strResultMessage = vbNullString

    Set oCandidate = pCandidate
    Set oTarget = pTarget

    'Set candidate and Target
    oCPSProfileHelper.Candidate = oCandidate
    oCPSProfileHelper.Target = oTarget
    
    ReDim saCandCheckValues(UBound(strKeywordChecksColl)) As Variant
    ReDim saTgtCheckValues(UBound(strKeywordChecksColl)) As Variant
    ReDim saStatus(UBound(strKeywordChecksColl)) As Variant
    
    For j = LBound(strKeywordChecksColl) To UBound(strKeywordChecksColl)

        bAPISuccess = False

        Select Case strKeywordChecksColl(j)
            Case "Material"
                bAPISuccess = oCPSCommonHelper.GetRelatedObjects(CP_Material, oCndtMaterial, oTgtMaterial)

                If bAPISuccess Then
                    vCndtPropValue = oCndtMaterial.MaterialType + " | " + oCndtMaterial.MaterialGrade
                    vTgtPropValue = oTgtMaterial.MaterialType + " | " + oTgtMaterial.MaterialGrade

                    If oCndtMaterial.MaterialType <> oTgtMaterial.MaterialType Or _
                        oCndtMaterial.MaterialGrade <> oTgtMaterial.MaterialGrade Then
                        bCommon = False
                    End If
                End If
                
            Case "Section"
                bAPISuccess = oCPSCommonHelper.GetRelatedObjects(CP_CatalogDefinition, oCndtXSection, oTgtXSection)

                If Not oCndtXSection Is Nothing And Not oTgtXSection Is Nothing And bAPISuccess Then
                          
                    bAPISuccess = False
                    bAPISuccess = oCPSCommonHelper.GetRelatedObjectPropertyValues(CP_CatalogDefinition, "SectionName", vCndtPropValue, vTgtPropValue)
                    
                    If bAPISuccess = False Or oCndtXSection.Type <> oTgtXSection.Type Or CStr(vCndtPropValue) <> CStr(vTgtPropValue) Then
                       bCommon = False
                    End If
                        
                    vCndtPropValue = oCndtXSection.Type + " | " + CStr(vCndtPropValue)
                    vTgtPropValue = oTgtXSection.Type + " | " + CStr(vTgtPropValue)
                End If

            Case "NamingCategory"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("NamingCategory", vCndtPropValue, vTgtPropValue) ',"IJPlate")

                If bAPISuccess Then
                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "ProfileType"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("pType", vCndtPropValue, vTgtPropValue) ',"IJStiffener")

                If bAPISuccess Then
                    eCndtProfileType = vCndtPropValue
                    eTgtProfileType = vTgtPropValue

                    If eCndtProfileType <> eTgtProfileType Then
                        bCommon = False
                    End If
                End If

            Case "Curved"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("Curved", vCndtPropValue, vTgtPropValue) ',"IJProfilePart")

                If bAPISuccess Then
                    eCndtCurvature = vCndtPropValue
                    eTgtCurvature = vTgtPropValue

                    If eCndtCurvature <> eTgtCurvature Then
                        bCommon = False
                    End If
                End If

            'double comparison
            Case "Area"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("Area", vCndtPropValue, vTgtPropValue) ',"IJProfilePart")

                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dAreaTolerance Then
                        bCommon = False
                    End If
                End If

            'double comparison
            Case "Length"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("ProfileLength", vCndtPropValue, vTgtPropValue) ',"IJProfilePart")

                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dDistanceTolerance Then
                        bCommon = False
                    End If
                End If

            Case "Twisted"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("IsTwisted", vCndtPropValue, vTgtPropValue) ',"IJProfilePart")

                If bAPISuccess Then
                    If CBool(vCndtPropValue) <> CBool(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            'double comparison
            Case "DryWeight"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("DryWeight", vCndtPropValue, vTgtPropValue) ',"IJWeightCG")

                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dWeightTolerance Then
                        bCommon = False
                    End If
                End If

            'double comparison
            Case "WetWeight"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("WetWeight", vCndtPropValue, vTgtPropValue) ',"IJWeightCG")

                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dWeightTolerance Then
                        bCommon = False
                    End If
                End If

            Case "BuildMethod"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("BuildMethod", vCndtPropValue, vTgtPropValue) ',"IJUAAssemblyChild")

                If bAPISuccess Then
                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "SlotConnectivity"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("SlotConnectivity", vCndtPropValue, vTgtPropValue) ',"IJUAAssemblyChild")

                If bAPISuccess Then
                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "StageCode"
                bAPISuccess = oCPSCommonHelper.GetRelatedObjectPropertyValues(CP_ProductionRouting, "StageCode", vCndtPropValue, vTgtPropValue) ', "IJPlnProductionRouting")

                If bAPISuccess Then
                    If CStr(vCndtPropValue) <> CStr(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "WorkCenter"
                bAPISuccess = oCPSCommonHelper.GetRelatedObjectPropertyValues(CP_ProductionRouting, "WorkCenter", vCndtPropValue, vTgtPropValue) ', "IJPlnProductionRouting")

                If bAPISuccess Then
                    If CStr(vCndtPropValue) <> CStr(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

           Case "Actions"
                bAPISuccess = oCPSCommonHelper.GetRelatedObjects(CP_ProductionRouting, oCndtProdRouting, oTgtProdRouting)

                If bAPISuccess Then 'And Not oCndtProdRouting Is Nothing Then ' And Not oTgtProdRouting Is Nothing Then
                    If Not oCndtProdRouting Is Nothing Then
                        Set oCndtProdRtgActions = GetRelatedObjects(oCndtProdRouting, IID_IJPlnProductionRouting, "RoutingAction")
                        vCndtPropValue = oCndtProdRtgActions.Count
                    Else
                        vCndtPropValue = ""
                    End If

                    If Not oTgtProdRouting Is Nothing Then
                        Set oTgtProdRtgActions = GetRelatedObjects(oTgtProdRouting, IID_IJPlnProductionRouting, "RoutingAction")
                        vTgtPropValue = oTgtProdRtgActions.Count
                    Else
                        vTgtPropValue = ""
                    End If

                    If CStr(vCndtPropValue) <> CStr(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

'            Case "Action"   'Need more inputs on this


            Case "Direction"
                bAPISuccess = oCPSProfileHelper.GetPrimaryOrientationSides(eCndtPrimOrientation, eTgtPrimOrientation)

                vCndtPropValue = CLng(eCndtPrimOrientation)
                vTgtPropValue = CLng(eTgtPrimOrientation)

                If bAPISuccess Then
                    If eCndtPrimOrientation <> eTgtPrimOrientation Then
                        bCommon = False
                    End If
                End If

            Case "PrimaryOrientation"
                bAPISuccess = oCPSProfileHelper.GetPrimaryOrientationVectors(oCndtOrientVector, oTgtOrientVector)

                oCndtOrientVector.Get dCndt_X, dCndt_Y, dCndt_Z
                oTgtOrientVector.Get dTgt_X, dTgt_Y, dTgt_Z

                If bAPISuccess Then
                    If Abs(dCndt_X - dTgt_X) > m_dDistanceTolerance Or Abs(dCndt_Y - dTgt_Y) > m_dDistanceTolerance Or Abs(dCndt_Z - dTgt_Z) > m_dDistanceTolerance Then
                        bCommon = False
                    End If
                End If

                vCndtPropValue = CStr(dCndt_X) + "/" + CStr(dCndt_Y) + "/" + CStr(dCndt_Z)
                vTgtPropValue = CStr(dTgt_X) + "/" + CStr(dTgt_Y) + "/" + CStr(dTgt_Z)

            Case "SecondaryOrientation"
                bAPISuccess = oCPSProfileHelper.GetSecondaryOrientationVectors(oCndtOrientVector, oTgtOrientVector)

                oCndtOrientVector.Get dCndt_X, dCndt_Y, dCndt_Z
                oTgtOrientVector.Get dTgt_X, dTgt_Y, dTgt_Z

                If bAPISuccess Then
                    If Abs(dCndt_X - dTgt_X) > m_dDistanceTolerance Or Abs(dCndt_Y - dTgt_Y) > m_dDistanceTolerance Or Abs(dCndt_Z - dTgt_Z) > m_dDistanceTolerance Then
                        bCommon = False
                    End If
                End If

                vCndtPropValue = CStr(dCndt_X) + "/" + CStr(dCndt_Y) + "/" + CStr(dCndt_Z)
                vTgtPropValue = CStr(dTgt_X) + "/" + CStr(dTgt_Y) + "/" + CStr(dTgt_Z)

            Case "CrossSectionDimensions"
                Dim dCndtFlangeLength       As Double
                Dim dTgtFlangeLength        As Double
                Dim dCndtWebLength          As Double
                Dim dTgtWebLength           As Double
                Dim dCndtFlangeThickness    As Double
                Dim dTgtFlangeThickness     As Double
                Dim dCndtWebThickness       As Double
                Dim dTgtWebThickness        As Double
                Dim Dcv                     As Variant
                Dim Dtv                     As Variant

                bAPISuccess = oCPSCommonHelper.GetRelatedObjectPropertyValues(CP_CatalogDefinition, "FlangeLength", Dcv, Dtv) ' dCndtFlangeLength, dTgtFlangeLength)

                dCndtFlangeLength = Dcv
                dTgtFlangeLength = Dtv

                If bAPISuccess Then
                    If Abs(dCndtFlangeLength - dTgtFlangeLength) > m_dDistanceTolerance Then
                        bCommon = False
                    End If

                    vCndtPropValue = CStr(dCndtFlangeLength)
                    vTgtPropValue = CStr(dTgtFlangeLength)
                End If

                bAPISuccess = oCPSCommonHelper.GetRelatedObjectPropertyValues(CP_CatalogDefinition, "WebLength", Dcv, Dtv) ' dCndtWebLength, dTgtWebLength)

                dCndtWebLength = Dcv
                dTgtWebLength = Dtv

                If bAPISuccess Then
                    If Abs(dCndtWebLength - dTgtWebLength) > m_dDistanceTolerance Then
                        bCommon = False
                    End If

                    vCndtPropValue = vCndtPropValue + "X " + CStr(dCndtWebLength)
                    vTgtPropValue = vTgtPropValue + "X " + CStr(dTgtWebLength)
                End If

                bAPISuccess = oCPSCommonHelper.GetRelatedObjectPropertyValues(CP_CatalogDefinition, "FlangeThickness", Dcv, Dtv) ' dCndtFlangeThickness, dTgtFlangeThickness)

                dCndtFlangeThickness = Dcv
                dTgtFlangeThickness = Dtv

                If bAPISuccess Then
                    If Abs(dCndtFlangeThickness - dTgtFlangeThickness) > m_dDistanceTolerance Then
                        bCommon = False
                    End If

                    vCndtPropValue = vCndtPropValue + "X " + CStr(dCndtFlangeThickness)
                    vTgtPropValue = vTgtPropValue + "X " + CStr(dTgtFlangeThickness)
                End If

                bAPISuccess = oCPSCommonHelper.GetRelatedObjectPropertyValues(CP_CatalogDefinition, "WebThickness", Dcv, Dtv) ' dCndtWebThickness, dTgtWebThickness)

                dCndtWebThickness = Dcv
                dTgtWebThickness = Dtv

                If bAPISuccess Then
                    If Abs(dCndtWebThickness - dTgtWebThickness) > m_dDistanceTolerance Then
                        bCommon = False
                    End If

                    vCndtPropValue = vCndtPropValue + "X " + CStr(dCndtWebThickness)
                    vTgtPropValue = vTgtPropValue + "X " + CStr(dTgtWebThickness)
                End If

            'double comparison
            Case "SurfaceArea"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("EstimatedSurfaceArea", vCndtPropValue, vTgtPropValue) ',"IJProfile")

                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dAreaTolerance Then
                        bCommon = False
                    End If
                End If

            Case "LandingCurveLength"
                bAPISuccess = oCPSProfileHelper.GetProfileLandingCurveLengths(dCndtLndCrvLength, dTgtLndCrvLength)

                If bAPISuccess Then
                    vCndtPropValue = dCndtLndCrvLength
                    vTgtPropValue = dTgtLndCrvLength

                    If Abs(dCndtLndCrvLength - dTgtLndCrvLength) > m_dDistanceTolerance Then
                        bCommon = False
                    End If
                End If

            Case "BoundingBox"  'Heard about a new CORE API to get this
                Dim dCndtLength     As Double
                Dim dTgtLength      As Double
                Dim dCndtWidth      As Double
                Dim dTgtWidth       As Double
                Dim dCndtThickness  As Double
                Dim dTgtThickness   As Double

                bAPISuccess = oCPSCommonHelper.GetBoundingBoxDimensions(dCndtLength, dCndtWidth, dCndtThickness, dTgtLength, dTgtWidth, dTgtThickness)

                If bAPISuccess Then
                    If Abs((dCndtLength * dCndtWidth * dCndtThickness) - (dTgtLength * dTgtWidth * dTgtThickness)) > m_dVolumeTolerance Then
                        bCommon = False
                    End If
                End If
                
                Dim DecimalNumber() As String
                DecimalNumber = Split(m_dVolumeTolerance, GetDecimalSeparator)
                vCndtPropValue = CStr(FormatNumber(dCndtLength, Len(DecimalNumber(1)))) + "; " + CStr(FormatNumber(dCndtWidth, Len(DecimalNumber(1)))) + "; " + CStr(FormatNumber(dCndtThickness, Len(DecimalNumber(1))))
                vTgtPropValue = CStr(FormatNumber(dTgtLength, Len(DecimalNumber(1)))) + "; " + CStr(FormatNumber(dTgtWidth, Len(DecimalNumber(1)))) + "; " + CStr(FormatNumber(dTgtThickness, Len(DecimalNumber(1))))
                
            Case "ProfileSolid"
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                bAPISuccess = True
                bCommon = bIsGeomSame

            Case "PCsCount"
                bAPISuccess = oCPSCommonHelper.GetConnectedObjects(AllPhysicalConnections, oCndtConnectedObjects, oTgtConnectedObjects)

                If bAPISuccess Then
                    vCndtPropValue = oCndtConnectedObjects.Count
                    vTgtPropValue = oTgtConnectedObjects.Count

                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "LateralConnections"
                bAPISuccess = oCPSCommonHelper.GetConnectedObjects(LateralPCs, oCndtConnectedObjects, oTgtConnectedObjects)

                If bAPISuccess Then
                    vCndtPropValue = oCndtConnectedObjects.Count
                    vTgtPropValue = oTgtConnectedObjects.Count

                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "MainConnections"
                bAPISuccess = oCPSCommonHelper.GetConnectedObjects(BaseAndOffsetPCs, oCndtConnectedObjects, oTgtConnectedObjects)

                If bAPISuccess Then
                    vCndtPropValue = oCndtConnectedObjects.Count
                    vTgtPropValue = oTgtConnectedObjects.Count

                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            'Bevels
            Case "ConnectionBevels"
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    If bPCsAlreadyMatched = False Then
                        bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(LateralPCs, oMatchedPCs)
                    Else
                        bAPISuccess = True
                    End If
                    
                    Dim bFlip As Boolean
                    bFlip = oCPSCommonHelper.Flipped

                    If bAPISuccess And oMatchedPCs.Count > 0 Then bCommon = CompareBevels(oCandidate, oTarget, oMatchedPCs, bFlip, strBevelInfo)
                    If bAPISuccess And oMatchedPCs.Count = 0 Then bCommon = True
                End If

            'MountingAngles
            Case "MountingAngles"
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    If bPCsAlreadyMatched = False Then
                        bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(AllPhysicalConnections, oMatchedPCs)
                    Else
                        bAPISuccess = True
                    End If

                    If bAPISuccess Then
                        bPCsAlreadyMatched = True
                        bCommon = ArePCMountingAnglesCommon(oMatchedPCs, oCPSCommonHelper.TransMatrix)
                    End If
                End If

            Case "FeaturesCount"  'Need new methods
                bAPISuccess = oCPSCommonHelper.GetConnectedObjects(AllFeatures, oCndtConnectedObjects, oTgtConnectedObjects)

                If bAPISuccess Then
                    vCndtPropValue = oCndtConnectedObjects.Count
                    vTgtPropValue = oTgtConnectedObjects.Count

                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False

                    End If

                End If

            Case "EdgeFeatureCount"  'Need new methods
                bAPISuccess = oCPSCommonHelper.GetConnectedObjects(EdgeFeatures, oCndtConnectedObjects, oTgtConnectedObjects)

                If bAPISuccess Then
                    vCndtPropValue = oCndtConnectedObjects.Count
                    vTgtPropValue = oTgtConnectedObjects.Count

                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "CornerFeatureCount" 'Need new methods
                bAPISuccess = oCPSCommonHelper.GetConnectedObjects(CornerFeatures, oCndtConnectedObjects, oTgtConnectedObjects)

                If bAPISuccess Then
                    vCndtPropValue = oCndtConnectedObjects.Count
                    vTgtPropValue = oTgtConnectedObjects.Count

                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "FreeEdgeTreatmentsCount" 'Need new methods
                bAPISuccess = oCPSCommonHelper.GetConnectedObjects(FreeEdgeTreatments, oCndtConnectedObjects, oTgtConnectedObjects)

                If bAPISuccess Then
                    vCndtPropValue = oCndtConnectedObjects.Count
                    vTgtPropValue = oTgtConnectedObjects.Count

                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "Features"  'Need new methods
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(AllFeatures, oConnectedObjPairs)
                    If bAPISuccess Then
                        bCommon = AreConnectedFeaturesCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix, False)
                    End If
                End If

            Case "EdgeFeatures"  'Need new methods
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(EdgeFeatures, oConnectedObjPairs)

                    If bAPISuccess Then
                        bCommon = AreConnectedFeaturesCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix, False)
                    End If

                End If

            Case "CornerFeatures" 'Need new methods
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(CornerFeatures, oConnectedObjPairs)

                    If bAPISuccess Then
                        bCommon = AreConnectedFeaturesCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix, False)
                    End If
                End If

            Case "FreeEdgeTreatments" 'Need new methods
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(FreeEdgeTreatments, oConnectedObjPairs)

                    If bAPISuccess Then
                        bCommon = AreConnectedFeaturesCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix)
                    End If
                End If

            Case "MarkingLines" 'Need new methods
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(MarkingLines, oConnectedObjPairs)

                    If bAPISuccess Then
                        bCommon = AreConnectedMfgObjectsCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix)
                    End If
                End If

            Case "Margins" 'Need new methods
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                 If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(Margins, oConnectedObjPairs)

                    If bAPISuccess Then
                        bCommon = AreConnectedMfgObjectsCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix)
                    End If
                 End If

            Case "Shrinkages" 'Need new methods
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(Shrinkages, oConnectedObjPairs)

                    If bAPISuccess Then
                        bCommon = AreConnectedMfgObjectsCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix)
                    End If
                End If

            Case "AssemblyShrinkages" 'Need new methods
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(AssemblyShrinkages, oConnectedObjPairs)

                    If bAPISuccess Then
                        bCommon = AreConnectedMfgObjectsCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix)
                    End If
                End If
        End Select

        If pManager Is Nothing And m_bStandardPartRule = False Then

            If strKeywordChecksColl(j) = "ProfileSolid" And bIsGeomSame = False Then
                Dim dCandidateVolume            As Double
                Dim dTargetVolume               As Double
                
                GetProfileVolumes pCandidate, pTarget, dCandidateVolume, dTargetVolume
            
                saCandCheckValues(j) = "Different. ( Volume: " & CStr(FormatNumber(dCandidateVolume, 6)) & " )"
                saTgtCheckValues(j) = "Different. ( Volume: " & CStr(FormatNumber(dTargetVolume, 6)) & " )"
                saStatus(j) = "Different"
            Else
                Dim arr() As String
                arr = Split(GetPropertyValues(oCandidate, vCndtPropValue, vTgtPropValue, bAPISuccess, bCommon, strKeywordChecksColl(j), oCPSCommonHelper, bIsGeomSame, oConnectedObjPairs, oMatchedPCs), "@")
                
                If strKeywordChecksColl(j) = "ConnectionBevels" Then
                    saCandCheckValues(j) = strBevelInfo
                    saTgtCheckValues(j) = strBevelInfo
                Else
                    saCandCheckValues(j) = arr(0)
                    saTgtCheckValues(j) = arr(1)
                End If
                saStatus(j) = CStr(arr(2))
            End If

        Else
            If bAPISuccess = False Then
                strResultMessage = "Middle Tier API failed while comparing " + strDisplayChecksColl(j)
                bCommon = False
                Exit For
            ElseIf bCommon = False Then
                Dim oCandidateNamedItem As IJNamedItem
                Set oCandidateNamedItem = pCandidate
            
                Dim oTargetNamedItem As IJNamedItem
                Set oTargetNamedItem = pTarget
                    
                If m_bGeometriesCompared = True And bIsGeomSame = False Then
                    ' It indicates that the geometry is not same
                    strResultMessage = "Geometry of Candidate part: " & oCandidateNamedItem.Name & " and Target Part: " & oTargetNamedItem.Name & " are not same"
                Else
                    strResultMessage = strDisplayChecksColl(j) + " of Candidate part: " & oCandidateNamedItem.Name & " and Target Part: " & oTargetNamedItem.Name & " are not same"
                End If
                
                Set oCandidateNamedItem = Nothing
                Set oTargetNamedItem = Nothing
                    
                Exit For
            End If
        End If

        vCndtPropValue = Null
        vTgtPropValue = Null
        bCommon = True
    Next j

    IJCommonEntityComparison_AreCommon = bCommon

    Set oCPSCommonHelper = Nothing
    Set oCPSProfileHelper = Nothing

Exit Function
ErrorHandler:
    IJCommonEntityComparison_AreCommon = False
    strResultMessage = "Unexpected error while comparing " + strDisplayChecksColl(j)
End Function

Private Sub IJCommonEntityComparison_SetPropertyValues(saPropertyValues() As String)
Const METHOD = "IJCommonEntityComparison_SetPropertyValues"
On Error GoTo ErrorHandler
    
    Dim i As Integer
    
    For i = LBound(saPropertyValues) To UBound(saPropertyValues) '2D Array of {tolerance Name, Value }
    
        Select Case saPropertyValues(i, 0)
            Case "DistanceTolerance"
                m_dDistanceTolerance = saPropertyValues(i, 1)
        
            Case "AreaTolerance"
                m_dAreaTolerance = saPropertyValues(i, 1)
        
            Case "VolumeTolerance"
                m_dVolumeTolerance = saPropertyValues(i, 1)
        
            Case "WeightTolerance"
                m_dWeightTolerance = saPropertyValues(i, 1)
                
            Case "AngleTolerance"
                m_dAngleTolerance = saPropertyValues(i, 1)
        End Select
    Next

Exit Sub
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Sub

Private Sub GetProfileVolumes(oCandidate As Object, oTarget As Object, dCandidateVolume As Double, dTargetVolume As Double)
Const METHOD = "GetProfileVolumes"
On Error GoTo ErrorHandler

    dCandidateVolume = 0
    dTargetVolume = 0
        
    Dim dAccuracyAchieved As Double, dLength As Double, dArea As Double, dVolume As Double
    Dim index As Long
        
    Dim oModelBody As IJDModelBody
    Set oModelBody = oCandidate
    
    oModelBody.GetDimMetrics m_dVolumeTolerance, dAccuracyAchieved, dLength, dArea, dVolume
    dCandidateVolume = dVolume

    Set oModelBody = Nothing
    Set oModelBody = oTarget
    
    oModelBody.GetDimMetrics m_dVolumeTolerance, dAccuracyAchieved, dLength, dArea, dVolume
    dTargetVolume = dVolume

Exit Sub
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Sub

Private Function IJStandardEntityComparison_IsStandardEntity(ByVal oCommonPartMember As Object, strDisplayChecksColl() As String, strKeywordChecksColl() As String, _
                                                                ByVal strStandardComparisonType As String, ByVal strStandardReferencePath As String, _
                                                                strStandardReferencePartName As String, strProcessPurpose As String) As Boolean
Const METHOD = "IJStandardEntityComparison_IsStandardEntity"
On Error GoTo ErrorHandler

    IJStandardEntityComparison_IsStandardEntity = False
        
    If strStandardComparisonType = "Model" Then
        ' get all the standard reference part collection from the input reference path
        Dim oStandardReferenceParts As IJDObjectCollection
        Set oStandardReferenceParts = GetStandardReferencePartsCollection(strStandardReferencePath)
        
        If oStandardReferenceParts Is Nothing Then
            Exit Function
        Else
            If oStandardReferenceParts.Count = 0 Then
                Exit Function
            End If
        End If
    
        Dim oStandardRefObj As Object
        
        ' Call AreCommon API for each standard reference part
        For Each oStandardRefObj In oStandardReferenceParts
            Dim strResultMsg  As String
            Dim saCandCheckValues() As Variant
            Dim saTgtCheckValues() As Variant
            Dim saStatus() As Variant
            Dim bCommon As Boolean
                        
            ' In case of Standard part rule execution, as Manager is not avaialable,
            ' needs to skip execution of AreCommon on first comparison failure so use m_bStandardPartRule as True
            m_bStandardPartRule = True
            bCommon = IJCommonEntityComparison_AreCommon(oCommonPartMember, oStandardRefObj, strDisplayChecksColl, strKeywordChecksColl, Nothing, _
                                                          strResultMsg, saCandCheckValues, saTgtCheckValues, saStatus)
                                                        
            ' If Common reference part is found, return refernce part name and TRUE to the caller
            If bCommon = True Then
                ' Return the standard part reference name
                Dim oStdPartNamedItem As IJNamedItem
                Set oStdPartNamedItem = oStandardRefObj
                strStandardReferencePartName = oStdPartNamedItem.Name
                
                ' ProcessPurpose codelist short description values => Common, StandardByModel, StandardByLookup
                strProcessPurpose = "StandardByModel"
                IJStandardEntityComparison_IsStandardEntity = True
                
                Exit For
            End If
        Next
    End If
    
    m_bStandardPartRule = False

Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Private Function ArePartGeometriesEqual(oCPSCommonHelper As IJPlnCompareHelperEx) As Boolean
Const METHOD = "ArePartGeometriesEqual"
On Error GoTo ErrorHandler

    m_bGeometriesCompared = True
    ArePartGeometriesEqual = False
    
    ' Setting the Property values to the helper
    ' Note that Common Part Service for Profiles needs only Distance Tolerance,
    ' Area Tolerance and Volume tolerance for geometry comparison
    oCPSCommonHelper.Properties(DistanceTolerance) = m_dDistanceTolerance
    oCPSCommonHelper.Properties(AreaTolerance) = m_dAreaTolerance
    oCPSCommonHelper.Properties(VolumeTolerance) = m_dVolumeTolerance
    
    ArePartGeometriesEqual = oCPSCommonHelper.IsGeometrySame(m_dDistanceTolerance)
    
Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function
