VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUMemberStd"
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

Const MODULE = "Built Up Member Standard: "

Const BU_MEMBER_STD_ASSEMBLY = 18

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
    m_oShrinkageHelper.ShrinkageRule = 106 ' Buitl Up Member Standard (StrMfgShrinkageRules.BUMemberStd) - will be useful to dump the rule name and rule progid
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
        
    'Compute Shrinkage paramaters for Std Memebr
    'Get Cross-sction and WeldingType so that we can check the look-up table
    Dim strSectionName As String
    Dim lWeldingType As Long
           
    'Get Cross-Scetion
    strSectionName = GetCrossSectionOfMember(oAssemblyBase)

    On Error Resume Next
    'Get Welding Type of Member
    lWeldingType = m_oShrinkageHelper.GetWeldingTypeOfBuiltUpMember(oAssemblyBase)
    On Error GoTo ErrorHandler

    '-----------------------------------------------------------
    'Log different assembly properties
    m_oShrinkageHelper.WriteAssemblyType lAssemblyType
    m_oShrinkageHelper.WriteMemberSectionName strSectionName
    m_oShrinkageHelper.WriteWeldType lWeldingType
    m_oShrinkageHelper.WriteBlankLine
    '-----------------------------------------------------------
    
    Dim oAlongWeldPort As IJPort
    Dim oNormalToWeldPort As IJPort
    m_oShrinkageHelper.GetPortsAlongWeldAndNormalToWeld oAssemblyBase, oAlongWeldPort, oNormalToWeldPort
    
    'Dump the vectors of the ports along welding and perpendicular to welding
    Dim oVectorAlongWeldingPort As IJDVector
    Dim oVectorPerpendicularWeldingPort As IJDVector
    
    Set oVectorAlongWeldingPort = GetVectorFromAxisOrPort(oAlongWeldPort)
    Set oVectorPerpendicularWeldingPort = GetVectorFromAxisOrPort(oNormalToWeldPort)
    
    '-----------------------------------------------------------
    'Log the different welding directions as vectors
    m_oShrinkageHelper.WriteWeldDirectionsForBuiltUps oAlongWeldPort, oNormalToWeldPort
    '-----------------------------------------------------------
    
    'Get the Direction and Values for Shrinkage from Look-Up JUAMfgShrinkageBUMemberCustom
    Dim lSecondaryAxis As Long
    Dim lPrimaryAxis As Long
    
    Dim dPrimaryVal As Double
    Dim dSecondaryVal As Double
    Dim dTertiaryVal As Double
    
    Dim bInvalidShrinkage As Boolean
    RunQueryOnBUMembeStdLookUp lAssemblyType, strSectionName, lWeldingType, lPrimaryAxis, lSecondaryAxis, dPrimaryVal, dSecondaryVal, dTertiaryVal, bInvalidShrinkage
    
    If bInvalidShrinkage = True Then
         ShrinkageType = ShrinkageUndefined
         GoTo CleanUp
    End If
    
    If Not oAlongWeldPort Is Nothing And Not oNormalToWeldPort Is Nothing Then
        ShrinkageType = ScalingType
    
        If lPrimaryAxis = 40 Then
            Set PrimaryAxis = oAlongWeldPort
            Set SecondaryAxis = oNormalToWeldPort
        ElseIf lPrimaryAxis = 41 Then
            Set PrimaryAxis = oNormalToWeldPort
            Set SecondaryAxis = oAlongWeldPort
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
    
    Exit Sub
ErrorHandler:
    m_oShrinkageHelper.WriteErrorDescription Err.Description
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7004, , "RULES")
    GoTo CleanUp
End Sub
Private Function GetCrossSectionOfMember(oAssemblyBase As IJAssemblyBase) As String
    Const METHOD = "GetCrossSectionOfMember"
    On Error GoTo ErrorHandler
    
    GetCrossSectionOfMember = vbNullString
     
    Dim oAssembly As IJAssembly
    Set oAssembly = oAssemblyBase
    
    Dim oChildren As IJDTargetObjectCol
    Set oChildren = oAssembly.GetChildren
    
    Dim iIndex As Integer
    For iIndex = 1 To oChildren.Count
        Dim oChild As Object
        Set oChild = oChildren.Item(iIndex)
        
        If TypeOf oChild Is IJPlatePart Then
            Dim oPlatePart As IJPlatePart
            Set oPlatePart = oChild
            Exit For
        End If
    Next
    
    If Not oPlatePart Is Nothing Then
        Dim oMemberSystem As Object
        Set oMemberSystem = GetParentMemberSystem(oPlatePart)
             
        Dim oISPSCross As ISPSCrossSection
        Set oISPSCross = oMemberSystem
    
        Dim oStrCoss As String
        oStrCoss = oISPSCross.SectionName
        
        GetCrossSectionOfMember = oStrCoss
    Else
        Exit Function
    End If
 
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub RunQueryOnBUMembeStdLookUp(lAssemblyType As Long, strSectionName As String, lWeldingType As Long, lPrimaryAxis As Long, lSecondaryAxis As Long, dPrimaryVal As Double, dSecondaryVal As Double, dTertiaryVal As Double, bNoRecordsFound As Boolean)
    Const METHOD = "RunQueryOnBUMembeStdLookUp"
    On Error GoTo ErrorHandler
                
    Dim strQuery As String
    Dim oQueryOutputValues() As Variant
    
    Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
    Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper

    strQuery = "SELECT PrimaryDirection, SecondaryDirection, PrimaryFactor, SecondaryFactor, TertiaryFactor FROM JUAMfgShrinkageBUMemberStd WHERE (AssemblyType = " + CStr(lAssemblyType) + ") AND (WeldingType = " + CStr(lWeldingType) + ") AND (SectionName = '" + strSectionName + "')"
        
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

Private Function GetParentMemberSystem(oPlatePart As Object) As Object
Const METHOD = "GetParentMemberSystem"
On Error GoTo ErrHandler

    Dim oChild   As IJSystemChild
    Dim oPlate   As Object

    Set oChild = oPlatePart
    
    Set oPlate = oChild.GetParent()
    
    If TypeOf oPlate Is ISPSDesignedMember Then ''TYpe -todo
        Set GetParentMemberSystem = oPlate
    Else
        Set GetParentMemberSystem = GetParentMemberSystem(oPlate)
    End If

CleanUp:
    Set oPlate = Nothing
    Set oChild = Nothing
    
    Exit Function
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
