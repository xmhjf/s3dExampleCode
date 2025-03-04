VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUProfileStd"
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
' 2012.JUNE.21    Raman     New desgin of the class
'*******************************************************************************
Option Explicit

Const MODULE = "Built Up Profile Standard Rule: "

Const BU_PROFILE_STD = 17

Private Enum eAssemblyDirection
    Assembly_Undefined = 0
    Assembly_Deck = 1
    Assembly_Transversal = 2
    Assembly_Longitudinal = 3
End Enum

Implements IJDShrinkageRule

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

    'If not an Assembly, return no shrinkage
    If Not TypeOf pDispObj Is IJProfilePart Then
        ShrinkageType = ShrinkageUndefined
        Exit Sub
    End If
            
    'Compute Shrinkage paramaters for Pre-SubBlock
    'Get Plate thickness,WeldingType so that we can check the look-up table
    Dim dWebThickness As Double
    Dim dWebHeight As Double
    Dim lWeldingType As Long
           
    'Get PlateThickness
    Dim oProfilePart As IJProfilePart
    Set oProfilePart = pDispObj
    
    If oProfilePart Is Nothing Then
        ShrinkageType = ShrinkageUndefined
        Exit Sub
    End If
    
    Dim strCrossSection As String
    GetCrossSectionOfProfile oProfilePart, strCrossSection
        
    'Get Major Welding Type and Direction
    Dim eStiffenerType As StructProfileType
    GetWeldingTypeAndTypeOfProfile oProfilePart, lWeldingType, eStiffenerType

    'Get the Direction and Values for Shrinkage from Look-Up JUAMfgShrinkageBUProfileStd
    Dim lSecondaryAxis As Long
    Dim lPrimaryAxis As Long
    
    Dim dPrimaryVal As Double
    Dim dSecondaryVal As Double
    Dim dTertiaryVal As Double
    
    Dim bInvalidShrinkage As Boolean
    RunQueryOnBUProfileStdLookUp strCrossSection, lWeldingType, lPrimaryAxis, lSecondaryAxis, dPrimaryVal, dSecondaryVal, dTertiaryVal, bInvalidShrinkage
    
    If bInvalidShrinkage = True Then
         ShrinkageType = ShrinkageUndefined
         Exit Sub
    End If

    Dim oPrimaryPortAsFacePort As IJPort
    Dim oStiffenerPart As IJStiffenerPart
    Set oStiffenerPart = oProfilePart
    oStiffenerPart.GetMountingFacePort oPrimaryPortAsFacePort, vbNullString
    ShrinkageType = ScalingType

    Set PrimaryAxis = oPrimaryPortAsFacePort
    Set SecondaryAxis = Nothing

    PrimaryFactor = dPrimaryVal

    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7004, , "RULES")
End Sub

Private Sub GetCrossSectionOfProfile(oProfilePart As IJProfilePart, strCrossSection As String)
    Const METHOD = "GetCrossSectionOfProfile"
    On Error GoTo ErrorHandler
    
    Dim oStructDetailProfile As New StructDetailObjects.ProfilePart
    Set oStructDetailProfile.object = oProfilePart

    strCrossSection = oStructDetailProfile.SectionName

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub GetWeldingTypeAndTypeOfProfile(oProfilePart As IJProfilePart, lWeldingType As Long, eStiffenerDirType As StructProfileType)
    Const METHOD = "GetWeldingTypeAndTypeOfProfile"
    On Error GoTo ErrorHandler
    
    lWeldingType = 2 ' Default- double welding

    Dim oStructDetailProfile As New StructDetailObjects.ProfilePart
    Set oStructDetailProfile.object = oProfilePart

    eStiffenerDirType = oStructDetailProfile.ProfileType
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub RunQueryOnBUProfileStdLookUp(strCrossSection As String, lWeldingType As Long, lPrimaryAxis As Long, lSecondaryAxis As Long, dPrimaryVal As Double, dSecondaryVal As Double, dTertiaryVal As Double, bNoRecordsFound As Boolean)

    Const METHOD = "RunQueryOnBUProfileStdLookUp"
    On Error GoTo ErrorHandler
    
    Dim strQuery As String
    Dim oQueryOutputValues() As Variant
    
    Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
    Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper
        
    strQuery = "SELECT PrimaryDirection, SecondaryDirection, PrimaryFactor, SecondaryFactor FROM JUAMfgShrinkageBUProfileStd WHERE (SectionName = '" + strCrossSection + "')"
        
    On Error Resume Next
    oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)
    
    If (UBound(oQueryOutputValues) > 0) Then
        lPrimaryAxis = oQueryOutputValues(0)
        lSecondaryAxis = oQueryOutputValues(1)
        dPrimaryVal = oQueryOutputValues(2)
        dSecondaryVal = oQueryOutputValues(3)
        dTertiaryVal = oQueryOutputValues(4)
    Else
        bNoRecordsFound = True
    End If
    
    Exit Sub
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
