VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation. All rights reserved.
'
'  Project     : SharedContent\Src\Planning\Rules\CommonEntity\ComparisonRule\CommonEntityComparison.vbp
'  File        : PlateRule.cls
'
'  Description : Common Part Plate comparison rule
'
'  History     :
'   9th Jan 2012      Devi Kishore Adiraju     Initial creation
'**************************************************************************************

Option Explicit

Private m_strLogFile                    As String
Private m_oStream                       As TextStream

Private Const IID_IJAssemblyChild As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
Private Const IID_IJPlnProductionRouting As String = "{E9B1D9F6-C687-40BA-85E2-311534F4782A}"

Implements IJCommonEntityComparison
Implements IJStandardEntityComparison

Private Sub Class_Initialize()

    sSOURCEFILE = "PlateRule.cls"
    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject(ERRORPROGID)
    End If
    
    'How should tolerances be read from the XML?
    'XML should be sent as an input to this class

    m_dDistanceTolerance = 0.001
    m_dThicknessTolerance = 0.000001
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
    Dim oCandidate                  As IJPlatePart
    Dim oTarget                     As IJPlatePart
    Dim bCommon                     As Boolean
    Dim oCndtMaterial               As IJDMaterial
    Dim oTgtMaterial                As IJDMaterial
    Dim oCndtDefinition             As Object
    Dim oTgtDefinition              As Object
    Dim oCndtSmartItem              As IJSmartItem
    Dim oTgtSmartItem               As IJSmartItem
    Dim oCndtDefCatalogPart         As IJDPart
    Dim oTgtDefCatalogPart          As IJDPart
    Dim eCndtPlateType              As StructPlateType
    Dim eTgtPlateType               As StructPlateType
    Dim eCndtTightness              As StructPlateTightness
    Dim eTgtTightness               As StructPlateTightness
    Dim eCndtCurvature              As PlateCurvature
    Dim eTgtCurvature               As PlateCurvature
    Dim oCndtBoundaryCurve          As IJDModelBody
    Dim oTgtBoundaryCurve           As IJDModelBody
    Dim oConnectedObjPairs          As IJElements
    Dim oCndtConnectedObjects       As IJElements
    Dim oTgtConnectedObjects        As IJElements
    Dim oCndtProdRouting            As Object
    Dim oTgtProdRouting             As Object
    Dim oCndtProdRtgActions         As IJElements
    Dim oTgtProdRtgActions          As IJElements
    Dim oMatchedPCs                 As IJElements
    Dim bPCsAlreadyMatched          As Boolean
    
    Dim vCndtPropValue              As Variant
    Dim vTgtPropValue               As Variant
    Dim bAPISuccess                 As Boolean
    Dim bIsGeomSame                 As Boolean
    
    'Create a Common Part Plate Helper and Common Helper
    Dim oCPSCommonHelper            As IJPlnCompareHelperEx
    Dim oCPSPlateHelper             As IJPlnPlateHelper
    Dim strBevelInfo                As String
    
    Set oCPSPlateHelper = New CPlnPlateHelper
    Set oCPSCommonHelper = oCPSPlateHelper
    
    Set oCandidate = pCandidate
    Set oTarget = pTarget
    
    'Set candidate and Target
    oCPSPlateHelper.Candidate = oCandidate
    oCPSPlateHelper.Target = oTarget
    
    'By default the comparison returns true
    bCommon = True
    m_bGeometriesCompared = False   'Indicates that the geometries are not compared yet.
    strResultMessage = vbNullString

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
                    
            'double comparison of plate thickness
             Case "Thickness"
                Dim oPlate As IJPlate
                Set oPlate = oCandidate
                vCndtPropValue = oPlate.thickness
                
                Set oPlate = Nothing
                Set oPlate = oTarget
                vTgtPropValue = oPlate.thickness
                
                bAPISuccess = True
                
                ' Use Thickness tolerance for Thickness comparison (as thickness values are very small)
                If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dThicknessTolerance Then
                    bCommon = False
                End If
                
            'Check catalog item
            Case "Definition"
                bAPISuccess = oCPSCommonHelper.GetRelatedObjects(CP_CatalogDefinition, oCndtDefinition, oTgtDefinition)
                
                If Not oCndtDefinition Is Nothing And Not oTgtDefinition Is Nothing And bAPISuccess Then
                    If TypeOf oCndtDefinition Is IJSmartItem And TypeOf oTgtDefinition Is IJSmartItem Then
                        Set oCndtSmartItem = oCndtDefinition
                        Set oTgtSmartItem = oTgtDefinition

                        vCndtPropValue = oCndtSmartItem.Name
                        vTgtPropValue = oTgtSmartItem.Name
                    ElseIf TypeOf oCndtDefinition Is IJDPart And TypeOf oTgtDefinition Is IJDPart Then
                        Set oCndtDefCatalogPart = oCndtDefinition
                        Set oTgtDefCatalogPart = oTgtDefinition
                        
                        vCndtPropValue = oCndtDefCatalogPart.PartNumber
                        vTgtPropValue = oCndtDefCatalogPart.PartNumber
                    End If
                    
                    If CStr(vCndtPropValue) <> CStr(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If
            
            Case "NamingCategory"
                 bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("NamingCategory", vCndtPropValue, vTgtPropValue) ',"IJPlate")

                 If bAPISuccess Then
                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                 End If
            
            Case "PlateType"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("PlateType", vCndtPropValue, vTgtPropValue) ',"IJPlate")
                
                If bAPISuccess Then
                    eCndtPlateType = vCndtPropValue
                    eTgtPlateType = vTgtPropValue
                    
                    If eCndtPlateType <> eTgtPlateType Then
                        bCommon = False
                    End If
                End If
                                
            Case "Tightness"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("Tightness", vCndtPropValue, vTgtPropValue) ',"IJPlate")
                
                If bAPISuccess Then
                    eCndtTightness = vCndtPropValue
                    eTgtTightness = vTgtPropValue
                    
                    If eCndtTightness <> eTgtTightness Then
                        bCommon = False
                    End If
                End If
            
            Case "Curved"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("Curved", vCndtPropValue, vTgtPropValue) ',"IJPlatePart")
                               
                If bAPISuccess Then
                    eCndtCurvature = vCndtPropValue
                    eTgtCurvature = vTgtPropValue
                    
                    If eCndtCurvature <> eTgtCurvature Then
                        bCommon = False
                    End If
                End If
            
            'double comparison
            Case "Area"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("Area", vCndtPropValue, vTgtPropValue) ',"IJPlatePart")
                    
                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dAreaTolerance Then
                        bCommon = False
                    End If
                End If
                
            'double comparison
            Case "Length"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("PlateLength", vCndtPropValue, vTgtPropValue) ',"IJPlatePart")
                
                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dDistanceTolerance Then
                        bCommon = False
                    End If
                End If
                
            'double comparison
            Case "Width"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("PlateWidth", vCndtPropValue, vTgtPropValue) ',"IJPlatePart")
                
                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dDistanceTolerance Then
                        bCommon = False
                    End If
                End If
                
            'double comparison
            Case "DryWeight"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("DryWeight", vCndtPropValue, vTgtPropValue) ' "IJWeightCG")
                
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
                
                If bAPISuccess Then
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
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("fstDirection", vCndtPropValue, vTgtPropValue) ',"IJShpStrPartPosition")
                
                 vCndtPropValue = CLng(vCndtPropValue)
                 vTgtPropValue = CLng(vTgtPropValue)
                
                
                If bAPISuccess Then
                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If
            
            'double comparison
            Case "SurfaceArea"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("SurfaceArea", vCndtPropValue, vTgtPropValue) ',"IJPlate")
                
                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dAreaTolerance Then
                        bCommon = False
                    End If
                End If
            
            Case "TypeString"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("TypeString", vCndtPropValue, vTgtPropValue) ',"IJStructTypeString")
                
                If bAPISuccess Then
                    If CStr(vCndtPropValue) <> CStr(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If
            
            Case "BoundaryLength"
                bAPISuccess = oCPSPlateHelper.GetBoundaryCurves(oCndtBoundaryCurve, oTgtBoundaryCurve)
                
                If bAPISuccess Then
                    Dim dDummy1 As Double, dDummy2 As Double, dDummy3 As Double
                    Dim dCndtBndryLen As Double, dTgtBndryLen As Double
                    
                    oCndtBoundaryCurve.GetDimMetrics m_dDistanceTolerance, dDummy1, dCndtBndryLen, dDummy2, dDummy3
                    oTgtBoundaryCurve.GetDimMetrics m_dDistanceTolerance, dDummy1, dTgtBndryLen, dDummy2, dDummy3
                    
                    vCndtPropValue = dCndtBndryLen
                    vTgtPropValue = dTgtBndryLen
                    
                    If Abs(dCndtBndryLen - dTgtBndryLen) > m_dDistanceTolerance Then
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
                
            Case "Boundary"
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
                    
                    If bAPISuccess Then
                        bPCsAlreadyMatched = True
                    End If
                                          
                    Dim bFlip As Boolean
                    bFlip = oCPSPlateHelper.IsFlipped
                    
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
        
            If strKeywordChecksColl(j) = "Boundary" And bIsGeomSame = False Then
                Dim dCandidateLength            As Double
                Dim dTargetLength               As Double
            
                GetPlateOuterContourDimensions pCandidate, pTarget, dCandidateLength, dTargetLength
    
                saCandCheckValues(j) = "Different. ( Length: " & CStr(FormatNumber(dCandidateLength, 4)) & " )"
                saTgtCheckValues(j) = "Different. ( Length: " & CStr(FormatNumber(dTargetLength, 4)) & " )"
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
    Set oCPSPlateHelper = Nothing
    
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
                
            Case "ThicknessTolerance"
                m_dThicknessTolerance = saPropertyValues(i, 1)

            Case "AreaTolerance"
                m_dAreaTolerance = saPropertyValues(i, 1)

            Case "VolumeTolerance"
                m_dVolumeTolerance = saPropertyValues(i, 1)

            Case "WeightTolerance"
                m_dWeightTolerance = saPropertyValues(i, 1)
                
            Case "FlippingPart"
                m_bflippingpart = saPropertyValues(i, 1)
                
            Case "AngleTolerance"
                m_dAngleTolerance = saPropertyValues(i, 1)
        End Select
    Next

Exit Sub
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Sub

Private Sub GetPlateOuterContourDimensions(oCandidate As Object, oTarget As Object, dCandidateLength As Double, dTargetLength As Double)
Const METHOD = "GetPlateOuterContourDimensions"
On Error GoTo ErrorHandler

    dCandidateLength = 0
    dTargetLength = 0
    
    Dim oPlatePartSupport As IJPlatePartSupport
    Dim oPartSupport As IJPartSupport
    Dim oModelBody As IJDModelBody
    
    ' Compute candidate plate contour length
    Set oPlatePartSupport = New GSCADSDPartSupport.PlatePartSupport
    Set oPartSupport = oPlatePartSupport
    Set oPartSupport.Part = oCandidate
    
    Dim oBasePort As IJPort
    Dim oOffsetPort As IJPort
    
    oPlatePartSupport.GetSurfacePort PlateBaseSide, oBasePort
    oPlatePartSupport.GetSurfacePort PlateOffsetSide, oOffsetPort
    
    Dim oBasePortGeom As Object
    Dim oOffsetPortGeom As Object
    
    Set oBasePortGeom = oBasePort.Geometry
    Set oOffsetPortGeom = oOffsetPort.Geometry
        
    Dim oCPSPlateHelper  As IJPlnPlateHelper
    Set oCPSPlateHelper = New CPlnPlateHelper
    
    Dim oBaseContoursColl As IJElements
    Dim oOffsetContoursColl As IJElements
    
    oCPSPlateHelper.GetSurfaceBodyOuterContour oBasePortGeom, oBaseContoursColl
    oCPSPlateHelper.GetSurfaceBodyOuterContour oOffsetPortGeom, oOffsetContoursColl
    
    Dim dAccuracyAchieved As Double, dLength As Double, dArea As Double, dVolume As Double
    Dim index As Long
    
    For index = 1 To oBaseContoursColl.Count
        Set oModelBody = Nothing
        Set oModelBody = oBaseContoursColl.Item(index)
        
        oModelBody.GetDimMetrics m_dDistanceTolerance, dAccuracyAchieved, dLength, dArea, dVolume
        dCandidateLength = dCandidateLength + dLength
    Next
    
    For index = 1 To oOffsetContoursColl.Count
        Set oModelBody = Nothing
        Set oModelBody = oOffsetContoursColl.Item(index)
        
        oModelBody.GetDimMetrics m_dDistanceTolerance, dAccuracyAchieved, dLength, dArea, dVolume
        dCandidateLength = dCandidateLength + dLength
    Next
    
    ' Take average of base and offset contour lengths
    dCandidateLength = dCandidateLength / 2

    ' Compute target plate contour length
    Set oPartSupport.Part = oTarget
    
    Set oBasePort = Nothing
    Set oOffsetPort = Nothing
    
    oPlatePartSupport.GetSurfacePort PlateBaseSide, oBasePort
    oPlatePartSupport.GetSurfacePort PlateOffsetSide, oOffsetPort
    
    Set oBasePortGeom = Nothing
    Set oOffsetPortGeom = Nothing
    
    Set oBasePortGeom = oBasePort.Geometry
    Set oOffsetPortGeom = oOffsetPort.Geometry
    
    Set oBaseContoursColl = Nothing
    Set oOffsetContoursColl = Nothing
    
    oCPSPlateHelper.GetSurfaceBodyOuterContour oBasePortGeom, oBaseContoursColl
    oCPSPlateHelper.GetSurfaceBodyOuterContour oOffsetPortGeom, oOffsetContoursColl
    
    For index = 1 To oBaseContoursColl.Count
        Set oModelBody = Nothing
        Set oModelBody = oBaseContoursColl.Item(index)
        
        oModelBody.GetDimMetrics m_dDistanceTolerance, dAccuracyAchieved, dLength, dArea, dVolume
        dTargetLength = dTargetLength + dLength
    Next
    
    For index = 1 To oOffsetContoursColl.Count
        Set oModelBody = Nothing
        Set oModelBody = oOffsetContoursColl.Item(index)
        
        oModelBody.GetDimMetrics m_dDistanceTolerance, dAccuracyAchieved, dLength, dArea, dVolume
        dTargetLength = dTargetLength + dLength
    Next
    
    ' Take average of base and offset contour lengths
    dTargetLength = dTargetLength / 2

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
    ' Note that Common Part Service for Plates needs only Distance Tolerance,
    ' Area Tolerance and flag for part to be flipped for geometry comparison
    oCPSCommonHelper.Properties(DistanceTolerance) = m_dDistanceTolerance
    oCPSCommonHelper.Properties(AreaTolerance) = m_dAreaTolerance
    oCPSCommonHelper.Properties(FlippingPart) = m_bflippingpart
        
    ArePartGeometriesEqual = oCPSCommonHelper.IsGeometrySame(m_dDistanceTolerance)
    
Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function
