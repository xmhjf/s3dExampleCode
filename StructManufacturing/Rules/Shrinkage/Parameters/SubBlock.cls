VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgShrinkageRules
' Module: GrandBlockRule
'
' Description:  Determines the proposed settings for the GrandBlock type assembly
'
' Author:
'
' Comments:
' 2012.JUNE.21    Raman     Initial design of the class
'*******************************************************************************
Option Explicit

Const MODULE = "Sub Block: "
Const SUB_BLOCK = 6

Implements IJDShrinkageRule

Private m_oShrinkageHelper As MfgShrinkageHelper

Private Function IJDShrinkageRule_GetConnectedObjectsForPart(ByVal pDispObj As Object) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
    Const METHOD = "IJDShrinkageRule_GetConnectedObjectsForPart"
    On Error GoTo ErrorHandler
    
    'No implementation
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IJDShrinkageRule_GetDependentProfileShrinkageParameters(ByVal pDispObj As Object, ByVal pPlateColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlatePrimaryFactorColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlateSecondaryFactorColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlatePrimaryAxisColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlateSecondaryAxisColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, PrimaryFactor As Double)
    Const METHOD = "IJDShrinkageRule_GetDependentProfileShrinkageParameters"
    On Error GoTo ErrorHandler
        
    'No implementation
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Sub IJDShrinkageRule_GetShrinkageParameters(ByVal pDispObj As Object, ShrinkageType As StrMfgShrinkageType, PrimaryAxis As Object, PrimaryFactor As Double, SecondaryAxis As Object, SecondaryFactor As Double, Optional TertiaryAxis As Object, Optional TertiaryFactor As Double)
Const METHOD = "IJDShrinkageRule_GetShrinkageParameters"
On Error GoTo ErrorHandler
    
    Set m_oShrinkageHelper = New MfgShrinkageHelper
        
    'If not an Assembly, return no shrinkage
    If Not TypeOf pDispObj Is IJAssemblyBase Then
        ShrinkageType = ShrinkageUndefined
        Exit Sub
    End If
    
    'Set properties on shrinkage helper object
    m_oShrinkageHelper.ShrinkageRule = 101 ' Sub Block Rule (StrMfgShrinkageRules.SubBlock) - will be useful to dump the rule name and rule progid
    m_oShrinkageHelper.ShrinkageMode = AssemblyShrinkage
    m_oShrinkageHelper.ShrinkageParent = pDispObj
    
    'Start the logging
    m_oShrinkageHelper.WriteHeader
    
    'Log the rule log start time
    m_oShrinkageHelper.WriteTimeStamp
            
    'Log rule basic information
    m_oShrinkageHelper.WriteObjectRuleInfo
    m_oShrinkageHelper.WriteBlankLine
    
    Dim oAssemblyBase  As IJAssemblyBase
    Set oAssemblyBase = pDispObj
    
    'get the assembly type
    Dim lAssemblyType As Long
    lAssemblyType = oAssemblyBase.Type
        
    'Compute Shrinkage paramaters for Sub-Block
    'Get Plate thickness,WeldingType so that we can check the look-up table
    Dim dThickness As Double
    Dim lWeldingType As Long
           
    Dim oPlatePart As IJPlatePart
    Set oPlatePart = GetBasePlateofAssembly(oAssemblyBase)
   
    'log the base plate name
    m_oShrinkageHelper.WriteBasePlateName oPlatePart
    
    If oPlatePart Is Nothing Then ' no base plate - exit rule
        ShrinkageType = ShrinkageUndefined
        m_oShrinkageHelper.WriteExitRule
        GoTo CleanUp
        Exit Sub
    End If

    'Get PlateThickness
    GetThicknessOfBasePlate oPlatePart, dThickness

    'Get Major Welding Type and Direction
    Dim eAlongWelding As StructProfileType
    Dim eNormalToWelding As StructProfileType
    Dim eNormalToBasePlate As StructProfileType

    GetWeldingTypeAndDirection oPlatePart, lWeldingType, eAlongWelding, eNormalToWelding, eNormalToBasePlate

    '-----------------------------------------------------------
    'log the assembly properties like type, welding type and its base plate thickness
   
    m_oShrinkageHelper.WriteAssemblyType lAssemblyType
    m_oShrinkageHelper.WritePlateThickness dThickness
    m_oShrinkageHelper.WriteWeldType lWeldingType
    m_oShrinkageHelper.WriteBlankLine
    '-----------------------------------------------------------

    Dim oPlnIntHelper As IJDPlnIntHelper
    Set oPlnIntHelper = New CPlnIntHelper
    
    Dim dLength As Double
    Dim dWidth As Double
    Dim dHeight As Double
    
    oPlnIntHelper.GetDimension 2, oAssemblyBase, dLength, dWidth, dHeight 'using 2 as we want the range in global direction
    
    Dim dPrimaryRange As Double
    Dim dSecondaryRange As Double
    Dim dTertiaryRange As Double
    
    If eAlongWelding = sptLongitudinal Then
        dPrimaryRange = dLength
    ElseIf eAlongWelding = sptTransversal Then
        dPrimaryRange = dWidth
    ElseIf eAlongWelding = sptVertical Then
        dPrimaryRange = dHeight
    Else
        dPrimaryRange = dLength
    End If
    
    If eNormalToWelding = sptLongitudinal Then
        dSecondaryRange = dLength
    ElseIf eNormalToWelding = sptTransversal Then
        dSecondaryRange = dWidth
    ElseIf eNormalToWelding = sptVertical Then
        dSecondaryRange = dHeight
    Else
        dSecondaryRange = dWidth
    End If
    
    If eNormalToBasePlate = sptLongitudinal Then
        dTertiaryRange = dLength
    ElseIf eNormalToBasePlate = sptTransversal Then
        dTertiaryRange = dWidth
    ElseIf eNormalToBasePlate = sptVertical Then
        dTertiaryRange = dHeight
    Else
        dTertiaryRange = dHeight
    End If

    '-----------------------------------------------------------
    'log the assembly ranges
    m_oShrinkageHelper.WriteRangeInWeldingDirection dPrimaryRange
    m_oShrinkageHelper.WriteRangeInPerpendicularToWeldingDirection dSecondaryRange
    m_oShrinkageHelper.WriteRangeInNormalToBasePlateDirection dTertiaryRange
    '-----------------------------------------------------------

    Dim lSecondaryAxis As Long
    Dim lPrimaryAxis As Long

    Dim dPrimaryVal As Double
    Dim dSecondaryVal As Double
    Dim dTertiaryVal As Double

    Dim bInvalidShrinkage As Boolean
    RunQueryOnSubBlockLookUp lAssemblyType, dThickness, lWeldingType, dPrimaryRange, dSecondaryRange, dTertiaryRange, lPrimaryAxis, lSecondaryAxis, dPrimaryVal, dSecondaryVal, dTertiaryVal, bInvalidShrinkage

    If bInvalidShrinkage = True Then
         ShrinkageType = ShrinkageUndefined
         GoTo CleanUp
    End If

    Dim oFrameSystem As IHFrameSystem
    Set oFrameSystem = GetGlobalCoordinateSystem(oPlatePart)
    
    m_oShrinkageHelper.WriteFrameSystem oFrameSystem
    m_oShrinkageHelper.WriteBlankLine
    
    If Not oFrameSystem Is Nothing Then
        ShrinkageType = GlobalType
            
        If lPrimaryAxis = 20 Then
            Set PrimaryAxis = GetAxisFromEnum(eAlongWelding, oFrameSystem)
        ElseIf lPrimaryAxis = 21 Then
            Set PrimaryAxis = GetAxisFromEnum(eNormalToWelding, oFrameSystem)
        ElseIf lPrimaryAxis = 22 Then
            Set PrimaryAxis = GetAxisFromEnum(eNormalToBasePlate, oFrameSystem)
        End If
        
        If lSecondaryAxis = 20 Then
            Set SecondaryAxis = GetAxisFromEnum(eAlongWelding, oFrameSystem)
        ElseIf lSecondaryAxis = 21 Then
            Set SecondaryAxis = GetAxisFromEnum(eNormalToWelding, oFrameSystem)
        ElseIf lSecondaryAxis = 22 Then
            Set SecondaryAxis = GetAxisFromEnum(eNormalToBasePlate, oFrameSystem)
        End If
            
        PrimaryFactor = dPrimaryVal
        SecondaryFactor = dSecondaryVal
        TertiaryFactor = dTertiaryVal
    Else
        ShrinkageType = ShrinkageUndefined
        m_oShrinkageHelper.WriteExitRule
        GoTo CleanUp
    End If
    
    'Log the final data returned by the rule
    m_oShrinkageHelper.WriteRuleOutputData ShrinkageType, PrimaryAxis, SecondaryAxis, PrimaryFactor, SecondaryFactor, TertiaryFactor

CleanUp:
    ' Write the final log time
    m_oShrinkageHelper.WriteBlankLine
    m_oShrinkageHelper.WriteTimeStamp
    m_oShrinkageHelper.WriteHeader
    
    Set m_oShrinkageHelper = Nothing
    Set oPlatePart = Nothing
    
    Exit Sub
ErrorHandler:
    m_oShrinkageHelper.WriteErrorDescription Err.Description
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7004, , "RULES")
    GoTo CleanUp
End Sub
Private Sub GetThicknessOfBasePlate(oPlatePart As IJPlatePart, dThickness As Double)
    Const METHOD = "GetThicknessOfBasePlate"
    On Error GoTo ErrorHandler
    
    Dim oStrtDetailPlate As New StructDetailObjects.PlatePart
    Set oStrtDetailPlate.object = oPlatePart
    
    dThickness = oStrtDetailPlate.PlateThickness

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub GetWeldingTypeAndDirection(oPlatePart As IJPlatePart, lWeldingType As Long, eStiffenerDirType As StructProfileType, eNormalToWelding As StructProfileType, eNormalToBasePlate As StructProfileType)
    Const METHOD = "GetWeldingTypeAndDirection"
    On Error GoTo ErrorHandler

     'get the directions
    Dim oNumberL As Integer
    Dim oNumberT As Integer
    Dim oNumberV As Integer
    Dim oTempObj As Variant
    
    'Initialize values
    oNumberL = 0
    oNumberT = 0
    oNumberV = 0
            
    Dim oConnProfilesInAssembly As IJElements
    m_oShrinkageHelper.GetConnectedStiffenersInAssembly oPlatePart, oConnProfilesInAssembly
        
    For Each oTempObj In oConnProfilesInAssembly
        Dim oStiffner  As IJStiffener
        
        If TypeOf oTempObj Is IJStiffener Then
            Set oStiffner = oTempObj
            If oStiffner.pType = sptLongitudinal Then
                oNumberL = oNumberL + 1
            ElseIf oStiffner.pType = sptTransversal Then
                oNumberT = oNumberT + 1
            ElseIf oStiffner.pType = sptVertical Then
                oNumberV = oNumberV + 1
            End If
        End If
    Next
            
    Dim oSDPlatePart As New StructDetailObjects.PlatePart
    Set oSDPlatePart.object = oPlatePart
    
    Select Case oSDPlatePart.PlateType
        Case DeckPlate
            If oNumberL > oNumberT Then
                'deck stiffened along longitudinal
                eStiffenerDirType = sptLongitudinal
                eNormalToWelding = sptTransversal
                eNormalToBasePlate = sptVertical
            Else
                eStiffenerDirType = sptTransversal
                eNormalToWelding = sptLongitudinal
                eNormalToBasePlate = sptVertical
            End If
        Case TBulkheadPlate
            If oNumberT > oNumberV Then
                'Transversal BulkHead stiffened along Transeversal
                eStiffenerDirType = sptTransversal
                eNormalToWelding = sptVertical
                eNormalToBasePlate = sptLongitudinal
            Else
                'Transversal BulkHead stiffened along Vertical
                eStiffenerDirType = sptVertical
                eNormalToWelding = sptTransversal
                eNormalToBasePlate = sptLongitudinal
            End If
        Case LBulkheadPlate
            If oNumberL > oNumberV Then
                'Longitudinal Bulkhead stiffened along Longitudinal
                eStiffenerDirType = sptLongitudinal
                eNormalToWelding = sptVertical
                eNormalToBasePlate = sptTransversal
            Else
                'Longitudinal Bulkhead stiffened along Vertical
                eStiffenerDirType = sptVertical
                eNormalToWelding = sptLongitudinal
                eNormalToBasePlate = sptTransversal
            End If
    End Select
    
    If eStiffenerDirType = 0 Then
        eStiffenerDirType = sptTransversal
        eNormalToWelding = sptLongitudinal
        eNormalToBasePlate = sptVertical
    End If
    
    'to get welding type
    Dim oProfileInStiffDir As IJStiffener
    For Each oTempObj In oConnProfilesInAssembly
        If TypeOf oTempObj Is IJStiffener Then
            Set oStiffner = oTempObj
            If oStiffner.pType = eStiffenerDirType Then
                Set oProfileInStiffDir = oStiffner
            End If
        End If
    Next
    
    Dim retVal As Variant
    Dim retValType As EResultType
    
    'Retrieve welding type from PC
    If Not oProfileInStiffDir Is Nothing Then
        Dim oEntityHelper As MfgEntityHelper
        Set oEntityHelper = New MfgEntityHelper
        
        Dim oPhyConn As IJStructPhysicalConnection
        Set oPhyConn = oEntityHelper.GetPhysicalConnectionBetweenTwoObjects(oPlatePart, oProfileInStiffDir)
        
        Dim oStructDetailPhyConn As New StructDetailObjects.PhysicalConn
        Set oStructDetailPhyConn.object = oPhyConn
        
        oStructDetailPhyConn.GetCustomAttributeParamValue "IJUASmartWeld", "SingleSided", retVal, retValType
    End If
    
    If retVal = True Then
        lWeldingType = 1 'SingleSided
    Else
        lWeldingType = 2 'DoubleSided
'    Else
'        lWeldingType = 3 'Alternate
    End If
    
    '-----------------------------------------------------------
    ' log the number of profiles in each direction and the different welding directions
    m_oShrinkageHelper.WriteBlankLine
    m_oShrinkageHelper.WriteNumberOfProfiles oNumberL, oNumberT, oNumberV
    
    m_oShrinkageHelper.WriteBlankLine
    m_oShrinkageHelper.WriteWeldDirectionsForAssemblies GetStringFromStiffDirectionEnum(eStiffenerDirType), GetStringFromStiffDirectionEnum(eNormalToWelding), GetStringFromStiffDirectionEnum(eNormalToBasePlate)
    m_oShrinkageHelper.WriteBlankLine
    '-----------------------------------------------------------

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub RunQueryOnSubBlockLookUp(lAssemblyType As Long, dThickness As Double, lWeldingType As Long, dPrimaryRange As Double, _
                                        dSecondaryRange As Double, dTertiaryRange As Double, lPrimaryAxis As Long, lSecondaryAxis As Long, _
                                        dPrimaryVal As Double, dSecondaryVal As Double, dTertiaryVal As Double, bNoRecordsFound As Boolean)
    Const METHOD = "RunQueryOnSubBlockLookUp"
    On Error GoTo ErrorHandler
            
    Dim strQuery As String
    Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
    Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper
        
    Dim oQueryOutputValues() As Variant
    
    strQuery = "SELECT PrimaryDirection, SecondaryDirection, PrimaryFactor, SecondaryFactor, TertiaryFactor FROM JUAMfgShrinkageSubBlock WHERE (AssemblyType = " _
                                            + CStr(lAssemblyType) + ") AND (WeldingType = " + CStr(lWeldingType) + ") AND (RangePrimaryDirectionMin <= " _
                                            + CStr(dPrimaryRange) + ") AND (RangePrimaryDirectionMax > " + CStr(dPrimaryRange) + ") AND (RangeSecondaryDirectionMin <= " _
                                            + CStr(dSecondaryRange) + ") AND (RangeSecondaryDirectionMax > " + CStr(dSecondaryRange) + ") AND (RangeTertiaryDirectionMin <= " _
                                            + CStr(dTertiaryRange) + ") AND (RangeTertiaryDirectionMax > " + CStr(dTertiaryRange) + ") AND (PlateThicknessMin <= " _
                                            + CStr(dThickness) + ") AND (PlateThicknessMax > " + CStr(dThickness) + ")"
                                            
    
    On Error Resume Next
    oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)
    
    'logging the query
    m_oShrinkageHelper.WriteBlankLine
    m_oShrinkageHelper.WriteServiceTableQuery strQuery
    m_oShrinkageHelper.WriteBlankLine
    
    If (UBound(oQueryOutputValues) > 0) Then
        lPrimaryAxis = oQueryOutputValues(0)
        lSecondaryAxis = oQueryOutputValues(1)
        dPrimaryVal = oQueryOutputValues(2)
        dSecondaryVal = oQueryOutputValues(3)
        dTertiaryVal = oQueryOutputValues(4)
        
        'log the outputs from the table
        '-----------------------------------------------------------
        m_oShrinkageHelper.WriteServiceTableOutputs lPrimaryAxis, lSecondaryAxis, dPrimaryVal, dSecondaryVal, dTertiaryVal
        m_oShrinkageHelper.WriteBlankLine
        '-----------------------------------------------------------
    Else
        bNoRecordsFound = True
        'log that no outputs are found from the table
        '-----------------------------------------------------------
        m_oShrinkageHelper.WriteNoOutputsFound
        m_oShrinkageHelper.WriteExitRule
        m_oShrinkageHelper.WriteBlankLine
        '-----------------------------------------------------------
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function GetBasePlateofAssembly(oAssemblyBase As IJAssemblyBase) As IJPlatePart
    Const METHOD = "GetBasePlateofAssembly"
    On Error GoTo ErrorHandler
    
    Dim oAssemblyOrientation As IJAssemblyOrientation
    Set oAssemblyOrientation = oAssemblyBase
    
    Set GetBasePlateofAssembly = oAssemblyOrientation.GetBasePlate
    
    If GetBasePlateofAssembly Is Nothing Then
        Set GetBasePlateofAssembly = m_oShrinkageHelper.GetLargestPlateOfAssembly(oAssemblyBase)
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetAxisFromEnum(enumValue As StructProfileType, oFrameSystem As IHFrameSystem)
    If enumValue = sptLongitudinal Then
        Set GetAxisFromEnum = oFrameSystem.PrincipalXAxis
    ElseIf enumValue = sptTransversal Then
        Set GetAxisFromEnum = oFrameSystem.PrincipalYAxis
    ElseIf enumValue = sptVertical Then
        Set GetAxisFromEnum = oFrameSystem.PrincipalZAxis
    End If
End Function

