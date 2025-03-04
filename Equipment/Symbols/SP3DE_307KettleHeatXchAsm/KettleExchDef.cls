VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KettleExchDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
'   Copyright (C) 2007, Intergraph Corporation. All rights reserved.
'
'   File:      KettleExchDef.cls
'   Author:    VRK
'   Creation Date:  Monday,March 5 2007
'
'   Description:
'   This is E307 Kettle Heat Exchanger Equipment Assembly.
'   This Assembly has Five DatumPoints with Orientations X towards Up and Y along North, and
'   Two Rectangular Foundation Ports
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "SP3DE_307KettleHeatXchAsm:KettleExchDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors
Private m_avSymbolArrayOfInputs()   As Variant

Private m_oNorth                    As IJDVector
Private m_oEast                     As IJDVector
Private m_oElevation                As IJDVector
Private PI As Double

'Define Front End types
Private Const FrontEndType_A = 1
Private Const FrontEndType_B = 2
Private Const FrontEndType_C = 3
Private Const FrontEndType_D = 4
Private Const FrontEndType_N = 5
Private Const FrontEndType_Q = 6

'Define Vessel Head Types
Private Const EndType_2TO1 = 1
Private Const EndType_CONE = 2
Private Const EndType_DOME = 3
Private Const EndType_FnD = 4
Private Const EndType_FLAT = 5
Private Const EndType_FLGD = 6
Private Const EndType_HEMI = 7
Private Const EndType_NONE = 8
Private Const EndType_TORC = 9
Private Const EndType_TORS = 10

Private m_dFrontEndType As Long
Private m_dVesselHeadType As Long
Private m_dExchangerLength As Double
Private m_dExchangerNeckLength As Double
Private m_dExchangerNeckDiameter As Double
Private m_dExchangerTaperLength As Double
Private m_dExchangerDiameter As Double
Private m_dBotSupportCenFromPP As Double
Private m_dBottomSupportCentoCen As Double
Private m_dBottomSupportHeight As Double
Private m_dFrontEndFlangeDia As Double
Private m_dFrontEndFlangeTk1 As Double
Private m_dFrontEndLength1 As Double
Private m_dFrontEndLength2 As Double
Private m_dFrontEndFlangeTk2 As Double
Private m_dFrontEndFlangeTk3 As Double
Private m_dFrontEndDiameter As Double
Private m_dInsulationThickness As Double
Private m_dEndHeadConeTopDiameter As Double
Private m_dEndHeadConeHeight As Double
Private m_dEndHeadKnuckleRadius As Double
Private m_dEndHeadDomeradius As Double
Private m_dEndHeadFlangedThick1 As Double
Private m_dEndHeadFlangedThick2 As Double
Private m_dEndHeadSphericalRadius As Double
Private m_dFP1XBoltHole As Double
Private m_dFP1YBoltHole As Double
Private m_dFP2XBoltHole As Double
Private m_dFP2YBoltHole As Double
Private m_bComputeEqpComp As Boolean

'Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt

Private Sub Class_Initialize()
    Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors

    m_oEquipCADHelper.ProjectName = "SP3DE_307KettleHeatXchAsm"
    m_oEquipCADHelper.ClassName = "KettleExchDef"
    m_oEquipCADHelper.OccurrenceRootClass = orcEquipment
    LogCalls "Class_Initialize"
            
    Set m_oEast = New DVector
    m_oEast.x = 1
    m_oEast.y = 0
    m_oEast.z = 0
    
    Set m_oNorth = New DVector
    m_oNorth.x = 0
    m_oNorth.y = 1
    m_oNorth.z = 0
    
    Set m_oElevation = New DVector
    m_oElevation.x = 0
    m_oElevation.y = 0
    m_oElevation.z = 1
    
    PI = 4 * Atn(1)
End Sub

Private Sub Class_Terminate()
    LogCalls "Class_Terminate"
    Set m_oNorth = Nothing
    Set m_oEast = Nothing
    Set m_oElevation = Nothing
    Set m_oEditErrors = Nothing
    Set m_oEquipCADHelper = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation of IJDUserSymbolServices
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Progid should be unique
    IJDUserSymbolServices_GetDefinitionName = m_oEquipCADHelper.ProjectName & "." & m_oEquipCADHelper.ClassName
End Function

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                             ByVal defParams As Variant, _
                                                             ByVal pResourceMgr As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oSymbolDefinition As IJDSymbolDefinition
    Set oSymbolDefinition = m_oEquipCADHelper.InstanciateDefinition(CodeBase, defParams, pResourceMgr)
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oPropertyDescriptions As IJDPropertyDescriptions
    Dim oMemberDescriptions As IJDMemberDescriptions
    Dim oMemberDescription As IJDMemberDescription
    Dim oAggregatorDescription As IJDAggregatorDescription
    
    'Set up the aggregator
    Set oAggregatorDescription = oSymbolDefinition
    oAggregatorDescription.AggregatorClsid = m_oEquipCADHelper.OccurrenceRootClassGUID()
    oAggregatorDescription.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructEquipment"
    oAggregatorDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEquipment"
    oAggregatorDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEquipment"
    oAggregatorDescription.SetCMRemoveInputs imsCOOKIE_ID_USS_LIB, "CMRemoveInputsEquipment"
    
    'Add property to aggregator
    Set oPropertyDescriptions = oAggregatorDescription
    oPropertyDescriptions.RemoveAll
    oPropertyDescriptions.AddProperty "EquipmentProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEquipment", imsCOOKIE_ID_USS_LIB
    
    'Remove all the previous member descriptions
    Set oMemberDescriptions = oSymbolDefinition
    oMemberDescriptions.RemoveAll
    
    'Add your code here for the declaration of the Public Custom Methods used to manage new members
    'DP1
    'Add new member DP1 to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("DP1", 1, "CMConstructDP1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsDP1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructDP1"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalDP1"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseDP1"
    'Add properties for DP1
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "DP1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateDP1", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "DP1GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryDP1", imsCOOKIE_ID_USS_LIB
    
    'DP2
    'Add new member DP2 to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("DP2", 2, "CMConstructDP2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsDP2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructDP2"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalDP2"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseDP2"
    'Add properties for DP2
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "DP2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateDP2", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "DP2GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryDP2", imsCOOKIE_ID_USS_LIB
        
    'DP3
    'Add new member DP3 to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("DP3", 3, "CMConstructDP3", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsDP3"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructDP3"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalDP3"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseDP3"
    'Add properties for DP3
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "DP3Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateDP3", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "DP3GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryDP3", imsCOOKIE_ID_USS_LIB
        
    'DP4
    'Add new member DP4 to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("DP4", 4, "CMConstructDP4", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsDP4"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructDP4"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalDP4"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseDP4"
    'Add properties for DP4
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "DP4Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateDP4", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "DP4GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryDP4", imsCOOKIE_ID_USS_LIB
        
    'DP5
    'Add new member DP5 to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("DP5", 5, "CMConstructDP5", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsDP5"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructDP5"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalDP5"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseDP5"
    'Add properties for DP5
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "DP5Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateDP5", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "DP5GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryDP5", imsCOOKIE_ID_USS_LIB
            
    'Add new member(NozzleSTFndPort1) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleSTFndPort1", 6, "CMConstructNozzleSTFndPort1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleSTFndPort1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleSTFndPort1"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleSTFndPort1"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleSTFndPort1"
    'Add properties for (NozzleSTFndPort1)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleSTFndPort1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleSTFndPort1", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleSTFndPort1GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleSTFndPort1", imsCOOKIE_ID_USS_LIB

    'Add new member(NozzleSTFndPort2) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleSTFndPort2", 7, "CMConstructNozzleSTFndPort2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleSTFndPort2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleSTFndPort2"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleSTFndPort2"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleSTFndPort2"
    'Add properties for (NozzleSTFndPort2)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleSTFndPort2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleSTFndPort2", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleSTFndPort2GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleSTFndPort2", imsCOOKIE_ID_USS_LIB
        
    'Add new member(TyACDN) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("TyACDN", 8, "CMConstructTyACDN", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsTyACDN"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTyACDN"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalTyACDN"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseTyACDN"
    'Add properties for (TyACDN)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "TyACDNProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateTyACDN", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "TyACDNGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryTyACDN", imsCOOKIE_ID_USS_LIB

    'Add new member(TyB) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("TyB", 9, "CMConstructTyB", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsTyB"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTyB"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalTyB"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseTyB"
    'Add properties for (TyB)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "TyBProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateTyB", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "TyBGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryTyB", imsCOOKIE_ID_USS_LIB

    'Add new member(TyFrontQ) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("TyFrontQ", 10, "CMConstructTyFrontQ", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsTyFrontQ"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTyFrontQ"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalTyFrontQ"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseTyFrontQ"
    'Add properties for (TyFrontQ)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "TyFrontQProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateTyFrontQ", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "TyFrontQGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryTyFrontQ", imsCOOKIE_ID_USS_LIB
            
    'Add new member(EndType2T01) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("EndType2T01", 11, "CMConstructEndType2T01", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEndType2T01"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEndType2T01"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndType2T01"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndType2T01"
    'Add properties for (EndType2T01)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "EndType2T01Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateEndType2T01", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "EndType2T01GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEndType2T01", imsCOOKIE_ID_USS_LIB

    'Add new member(EndTypeCone) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("EndTypeCone", 12, "CMConstructEndTypeCone", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEndTypeCone"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEndTypeCone"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndTypeCone"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndTypeCone"
    'Add properties for (EndTypeCone)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "EndTypeConeProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEndTypeCone", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "EndTypeConeGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEndTypeCone", imsCOOKIE_ID_USS_LIB

    'Add new member(EndTypeDome) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("EndTypeDome", 13, "CMConstructEndTypeDome", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEndTypeDome"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEndTypeDome"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndTypeDome"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndTypeDome"
    'Add properties for (EndTypeDome)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "EndTypeDomeProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEndTypeDome", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "EndTypeDomeGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEndTypeDome", imsCOOKIE_ID_USS_LIB

    'Add new member(EndTypeFND) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("EndTypeFND", 14, "CMConstructEndTypeFND", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEndTypeFND"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEndTypeFND"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndTypeFND"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndTypeFND"
    'Add properties for (EndTypeFND)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "EndTypeFNDProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEndTypeFND", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "EndTypeFNDGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEndTypeFND", imsCOOKIE_ID_USS_LIB
    
    'Add new member(EndTypeFLGD) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("EndTypeFLGD", 15, "CMConstructEndTypeFLGD", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEndTypeFLGD"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEndTypeFLGD"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndTypeFLGD"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndTypeFLGD"
    'Add properties for (EndTypeFLGD)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "EndTypeFLGDProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEndTypeFLGD", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "EndTypeFLGDGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEndTypeFLGD", imsCOOKIE_ID_USS_LIB

    'Add new member(EndTypeHemi) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("EndTypeHemi", 16, "CMConstructEndTypeHemi", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEndTypeHemi"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEndTypeHemi"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndTypeHemi"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndTypeHemi"
    'Add properties for (EndTypeHemi)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "EndTypeHemiProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEndTypeHemi", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "EndTypeHemiGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEndTypeHemi", imsCOOKIE_ID_USS_LIB

    'Add new member(EndTypeTORC) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("EndTypeTORC", 17, "CMConstructEndTypeTORC", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEndTypeTORC"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEndTypeTORC"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndTypeTORC"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndTypeTORC"
    'Add properties for (EndTypeTORC)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "EndTypeTORCProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEndTypeTORC", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "EndTypeTORCGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEndTypeTORC", imsCOOKIE_ID_USS_LIB

    'Add new member(EndTypeTORS) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("EndTypeTORS", 18, "CMConstructEndTypeTORS", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEndTypeTORS"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEndTypeTORS"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndTypeTORS"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndTypeTORS"
    'Add properties for (EndTypeTORS)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "EndTypeTORSProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEndTypeTORS", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "EndTypeTORSGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEndTypeTORS", imsCOOKIE_ID_USS_LIB

    'Add new member(EndTypeFLAT) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("EndTypeFLAT", 19, "CMConstructEndTypeFLAT", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEndTypeFLAT"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEndTypeFLAT"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndTypeFLAT"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndTypeFLAT"
    'Add properties for (EndTypeFLAT)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "EndTypeFLATProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEndTypeFLAT", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "EndTypeFLATGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEndTypeFLAT", imsCOOKIE_ID_USS_LIB
            
    Set oAggregatorDescription = Nothing
    Set oMemberDescriptions = Nothing
    Set oMemberDescription = Nothing
    Set oPropertyDescriptions = Nothing
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
    'This method is not used by the CAD.
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    'Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
    IJDUserSymbolServices_EditOccurence = False
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation of IJEquipUserAttrMgmt
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJEquipUserAttrMgmt_OnAttributeChange(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As IJEquipAttrDescriptor, ByVal varNewAttrValue As Variant) As String
    Const METHOD = "IJEquipUserAttrMgmt_OnAttributeChange"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oMemberDescription As IJDMemberDescription
    Set oMemberDescription = m_oEquipCADHelper.GetMemberDescriptionFromChild(pIJDAttrs)
        
    Select Case oMemberDescription.Name
        Case "DP1", "DP2", "DP3", "DP4", "DP5", "NozzleSTFndPort1", "NozzleSTFndPort2", "TyACDN", "TyB", "TyFrontQ", _
                "EndType2T01", "EndTypeCone", "EndTypeDome", "EndTypeFND", "EndTypeFLGD", "EndTypeHemi", "EndTypeTORC", "EndTypeTORS", "EndTypeFLAT"
            Select Case UCase(pAttrToChange.InterfaceName)
                Case "IJDELETABLEMEMBER"
                    If UCase(pAttrToChange.AttrName) = "CANBEDELETED" Then
                        m_oEquipCADHelper.MakeMemberDeletable oMemberDescription, pIJDAttrs, CBool(varNewAttrValue)
                    End If
                Case Else
                    '
            End Select

        Case Else
            '
    End Select
    
    Set oMemberDescription = Nothing
    IJEquipUserAttrMgmt_OnAttributeChange = ""
    
    Exit Function
ErrorHandler:
    IJEquipUserAttrMgmt_OnAttributeChange = "ERROR"
    HandleError MODULE, METHOD
End Function

Private Function IJEquipUserAttrMgmt_OnPreCommit(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = "IJEquipUserAttrMgmt_OnPreCommit"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IJEquipUserAttrMgmt_OnPreCommit = "ERROR"
    IJEquipUserAttrMgmt_OnPreCommit = ""
   
    Exit Function
ErrorHandler:
    IJEquipUserAttrMgmt_OnPreCommit = "ERROR"
    HandleError MODULE, METHOD
End Function

Private Function IJEquipUserAttrMgmt_OnPreLoad(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = "IJEquipUserAttrMgmt_OnPreLoad"
    On Error GoTo ErrorHandler
    
    Dim oMemberDescription As IJDMemberDescription
    IJEquipUserAttrMgmt_OnPreLoad = "ERROR"
    Set oMemberDescription = m_oEquipCADHelper.GetMemberDescriptionFromChild(pIJDAttrs)
    Dim oAttrCollection As Collection
    Dim oAttributeDescriptor As IJEquipAttrDescriptor
    Dim m As Long
    
    Set oAttrCollection = CollAllDisplayedValues
    
    Select Case oMemberDescription.Name
        Case "DP1", "DP2", "DP3", "DP4", "DP5"
            For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                    Case "IJUADATUMSHAPE"
                        oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                    Case Else
                        '
                End Select
            Next
        Case "NozzleSTFndPort1", "NozzleSTFndPort2"
            For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                    Case "IJDELETABLEMEMBER"
                        If UCase(oAttributeDescriptor.AttrName) = "CANBEDELETED" Then
                            oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsChanged
                        End If
                        
                    Case "IJNOZZLEORIENTATION"
                        'The nozzle is placed by Point
                        oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                    Case Else
                    '
                End Select
            Next
        Case "TyACDN", "TyB", "TyFrontQ", "EndType2T01", "EndTypeCone", "EndTypeDome", "EndTypeFND", _
                "EndTypeFLGD", "EndTypeHemi", "EndTypeTORC", "EndTypeTORS", "EndTypeFLAT"
            For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                    Case "IJUAHEATEXCHANGERFRONTEND", "IJUAEXCHANGERCHANNEL", "IJUAHEATEXCHANGER", _
                            "IJUAEquipmentEnd", "IJUAVesselDiameter", "IJInsulationThickness"
                        oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                    Case Else
                        '
                End Select
            Next
            Set oAttrCollection = Nothing
        Case Else
        '
    End Select
    Set oAttributeDescriptor = Nothing
    Set oMemberDescription = Nothing
    Set oAttrCollection = Nothing
    IJEquipUserAttrMgmt_OnPreLoad = ""
    
    Exit Function
ErrorHandler:
    IJEquipUserAttrMgmt_OnPreLoad = "ERROR"
    HandleError MODULE, METHOD
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods of Equipment (Aggregator)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMConstructEquipment"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructEquipment"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMSetInputsEquipment"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    m_oEquipCADHelper.SetSmartItemAsInputToSymbol pAggregatorDescription
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMRemoveInputsEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMRemoveInputsEquipment"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEquipment(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEquipment"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub HandleError(sModule As String, sMethod As String)
    If Not m_oEditErrors Is Nothing Then
        m_oEditErrors.AddFromErr Err, "", sMethod, sModule
    End If
    
    Err.Raise Err.Number, Err.Source & " " & sMethod, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for DP1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructDP1(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oDatumShape As IJShape
    Dim oDesignEquipment As IJDesignEquipment
        
    'Create Datum Shape DP1
    Set oDatumShape = m_oEquipCADHelper.CreateShape(pMemberDescription, pResourceManager, "DatumShape 001", "DP1")
    If Not oDatumShape Is Nothing Then
        Set pObject = oDatumShape
        oDatumShape.RepresentationId = ReferenceGeometry

        Set oDesignEquipment = pMemberDescription.CAO
        oDesignEquipment.AddShape oDatumShape

        GetDimensionsFromSymbolArray oDesignEquipment
        PositionAndOrientDP1 oDesignEquipment, oDatumShape
    End If
    
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructDP1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsDP1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateDP1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryDP1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oEquipment As IJEquipment
    Dim oDatumShape As IJShape
    Set oDatumShape = oPropertyDescription.Object
    
    'Test if this DatumShape follows the symbol inputs
    If m_oEquipCADHelper.IsShapeFreeToTransform(oDatumShape) = False Then
        Set oEquipment = oPropertyDescription.CAO
        GetDimensionsFromSymbolArray oEquipment
        PositionAndOrientDP1 oEquipment, oDatumShape
    Else
        'Do nothing, the equipment will transform the shape for us
    End If
    
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    
    Exit Sub
ErrorHandler:
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalDP1(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseDP1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for DP2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructDP2(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oDatumShape As IJShape
    Dim oDesignEquipment As IJDesignEquipment
        
    'Create Datum Shape DP2
    Set oDatumShape = m_oEquipCADHelper.CreateShape(pMemberDescription, pResourceManager, "DatumShape 001", "DP2")
    If Not oDatumShape Is Nothing Then
        Set pObject = oDatumShape
        oDatumShape.RepresentationId = ReferenceGeometry

        Set oDesignEquipment = pMemberDescription.CAO
        oDesignEquipment.AddShape oDatumShape

        GetDimensionsFromSymbolArray oDesignEquipment
        PositionAndOrientDP2 oDesignEquipment, oDatumShape
    End If
    
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    
    Exit Sub
ErrorHandler:
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructDP2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsDP2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateDP2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryDP2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oEquipment As IJEquipment
    Dim oDatumShape As IJShape
    Set oDatumShape = oPropertyDescription.Object
    
    'Test if this DatumShape follows the symbol inputs
    If m_oEquipCADHelper.IsShapeFreeToTransform(oDatumShape) = False Then
        Set oEquipment = oPropertyDescription.CAO
        GetDimensionsFromSymbolArray oEquipment
        PositionAndOrientDP2 oEquipment, oDatumShape
    Else
        'Do nothing, the equipment will transform the shape for us
    End If
    
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    
    Exit Sub
ErrorHandler:
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalDP2(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseDP2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for DP3
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructDP3(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructDP3"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oDatumShape As IJShape
    Dim oDesignEquipment As IJDesignEquipment
    
    'Create Datum Shape DP3
    Set oDatumShape = m_oEquipCADHelper.CreateShape(pMemberDescription, pResourceManager, "DatumShape 001", "DP3")
    If Not oDatumShape Is Nothing Then
        Set pObject = oDatumShape
        oDatumShape.RepresentationId = ReferenceGeometry

        Set oDesignEquipment = pMemberDescription.CAO
        oDesignEquipment.AddShape oDatumShape

        GetDimensionsFromSymbolArray oDesignEquipment
        PositionAndOrientDP3 oDesignEquipment, oDatumShape
    End If
    
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    
    Exit Sub
ErrorHandler:
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructDP3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructDP3"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsDP3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsDP3"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateDP3(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateDP3"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryDP3(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryDP3"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEquipment As IJEquipment
    Dim oDatumShape As IJShape
    Set oDatumShape = oPropertyDescription.Object
    
    'Test if this DatumShape follows the symbol inputs
    If m_oEquipCADHelper.IsShapeFreeToTransform(oDatumShape) = False Then
        Set oEquipment = oPropertyDescription.CAO
        GetDimensionsFromSymbolArray oEquipment
        PositionAndOrientDP3 oEquipment, oDatumShape
    Else
        'Do nothing, the equipment will transform the shape for us
    End If
    
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    
    Exit Sub
ErrorHandler:
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalDP3(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalDP3"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseDP3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseDP3"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for DP4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructDP4(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructDP4"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oDatumShape As IJShape
    Dim oDesignEquipment As IJDesignEquipment
    
    'Create Datum Shape DP4
    Set oDatumShape = m_oEquipCADHelper.CreateShape(pMemberDescription, pResourceManager, "DatumShape 001", "DP4")
    If Not oDatumShape Is Nothing Then
        Set pObject = oDatumShape
        oDatumShape.RepresentationId = ReferenceGeometry

        Set oDesignEquipment = pMemberDescription.CAO
        oDesignEquipment.AddShape oDatumShape

        GetDimensionsFromSymbolArray oDesignEquipment
        PositionAndOrientDP4 oDesignEquipment, oDatumShape
    End If
    
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    
    Exit Sub
ErrorHandler:
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructDP4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructDP4"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsDP4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsDP4"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateDP4(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateDP4"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryDP4(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryDP4"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEquipment As IJEquipment
    Dim oDatumShape As IJShape
    Set oDatumShape = oPropertyDescription.Object
    
    'Test if this DatumShape follows the symbol inputs
    If m_oEquipCADHelper.IsShapeFreeToTransform(oDatumShape) = False Then
        Set oEquipment = oPropertyDescription.CAO
        GetDimensionsFromSymbolArray oEquipment
        PositionAndOrientDP4 oEquipment, oDatumShape
    Else
        'Do nothing, the equipment will transform the shape for us
    End If
    
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    
    Exit Sub
ErrorHandler:
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalDP4(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalDP4"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseDP4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseDP4"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for DP5
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructDP5(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructDP5"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oDatumShape As IJShape
    Dim oDesignEquipment As IJDesignEquipment
    
    'Create Datum Shape DP5
    Set oDatumShape = m_oEquipCADHelper.CreateShape(pMemberDescription, pResourceManager, "DatumShape 001", "DP5")
    If Not oDatumShape Is Nothing Then
        Set pObject = oDatumShape
        oDatumShape.RepresentationId = ReferenceGeometry

        Set oDesignEquipment = pMemberDescription.CAO
        oDesignEquipment.AddShape oDatumShape

        GetDimensionsFromSymbolArray oDesignEquipment
        PositionAndOrientDP5 oDesignEquipment, oDatumShape
    End If
    
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    
    Exit Sub
ErrorHandler:
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructDP5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructDP5"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsDP5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsDP5"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateDP5(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateDP5"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryDP5(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryDP5"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEquipment As IJEquipment
    Dim oDatumShape As IJShape
    Set oDatumShape = oPropertyDescription.Object
    
    'Test if this DatumShape follows the symbol inputs
    If m_oEquipCADHelper.IsShapeFreeToTransform(oDatumShape) = False Then
        Set oEquipment = oPropertyDescription.CAO
        GetDimensionsFromSymbolArray oEquipment
        PositionAndOrientDP5 oEquipment, oDatumShape
    Else
        'Do nothing, the equipment will transform the shape for us
    End If
    
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    
    Exit Sub
ErrorHandler:
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalDP5(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalDP5"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseDP5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseDP5"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
'The code of this method is specific to each Equipment as it converts the array of inputs in a set of global variables
Private Sub GetDimensionsFromSymbolArray(SmartOccurrence As IJSmartOccurrence)
    Const METHOD = "GetDimensionsFromSymbolArray"
    On Error GoTo ErrorHandler
    
    m_avSymbolArrayOfInputs = m_oEquipCADHelper.GetSymbolArrayOfInputs(SmartOccurrence)

    'Inputs,  from equipment symbol code
    m_dFrontEndType = m_avSymbolArrayOfInputs(2)
    m_dVesselHeadType = m_avSymbolArrayOfInputs(3)
    m_dExchangerLength = m_avSymbolArrayOfInputs(4)           'P1
    m_dExchangerNeckLength = m_avSymbolArrayOfInputs(5)       'P2
    m_dExchangerNeckDiameter = m_avSymbolArrayOfInputs(6)     'P3
    m_dExchangerTaperLength = m_avSymbolArrayOfInputs(7)      'P4
    m_dExchangerDiameter = m_avSymbolArrayOfInputs(8)         'P5
    m_dBotSupportCenFromPP = m_avSymbolArrayOfInputs(14)      'P10
    m_dBottomSupportCentoCen = m_avSymbolArrayOfInputs(15)    'P11
    m_dBottomSupportHeight = m_avSymbolArrayOfInputs(16)      'P14
    m_dFrontEndFlangeDia = m_avSymbolArrayOfInputs(19)        'P30
    m_dFrontEndFlangeTk1 = m_avSymbolArrayOfInputs(20)        'P31
    m_dFrontEndFlangeTk2 = m_avSymbolArrayOfInputs(21)        'P34
    m_dFrontEndFlangeTk3 = m_avSymbolArrayOfInputs(22)        'P35
    m_dFrontEndLength1 = m_avSymbolArrayOfInputs(23)          'P32
    m_dFrontEndLength2 = m_avSymbolArrayOfInputs(24)          'P33
    m_dFrontEndDiameter = m_avSymbolArrayOfInputs(25)         'P33 of type Q End
    m_dEndHeadConeTopDiameter = m_avSymbolArrayOfInputs(26)
    m_dEndHeadConeHeight = m_avSymbolArrayOfInputs(27)
    m_dEndHeadKnuckleRadius = m_avSymbolArrayOfInputs(28)
    m_dEndHeadDomeradius = m_avSymbolArrayOfInputs(29)
    m_dEndHeadFlangedThick1 = m_avSymbolArrayOfInputs(30)
    m_dEndHeadFlangedThick2 = m_avSymbolArrayOfInputs(31)
    m_dEndHeadSphericalRadius = m_avSymbolArrayOfInputs(32)
    m_dInsulationThickness = m_avSymbolArrayOfInputs(33)
    m_dFP1XBoltHole = m_avSymbolArrayOfInputs(34)
    m_dFP1YBoltHole = m_avSymbolArrayOfInputs(35)
    m_dFP2XBoltHole = m_avSymbolArrayOfInputs(36)
    m_dFP2YBoltHole = m_avSymbolArrayOfInputs(37)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub LogCalls(sMethod As String)
   m_oEquipCADHelper.LogError "Entering " & sMethod
End Sub

Private Sub PositionAndOrientDP1(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    
    oPosition.Set 0, 0, 0
    
    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oElevation
    
    Set oPosition = Nothing
End Sub

Private Sub PositionAndOrientDP2(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    
    oPosition.Set m_dFrontEndLength1 + m_dExchangerNeckLength + m_dExchangerTaperLength, 0, _
                    (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oElevation
    
    Set oPosition = Nothing
End Sub

Private Sub PositionAndOrientDP3(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    
    oPosition.Set m_dFrontEndLength1 + m_dExchangerLength, 0, _
                    (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
   
    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oElevation
    
    Set oPosition = Nothing
End Sub

Private Sub PositionAndOrientDP4(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    
    oPosition.Set m_dBotSupportCenFromPP, 0, _
        -(m_dBottomSupportHeight - (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2))

    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oElevation
    
    Set oPosition = Nothing
End Sub

Private Sub PositionAndOrientDP5(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    
    oPosition.Set m_dBotSupportCenFromPP + m_dBottomSupportCentoCen, 0, _
        -(m_dBottomSupportHeight - (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2))
    
    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oElevation
    
    Set oPosition = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleSTFndPort1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleSTFndPort1(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleSTFndPort1"
    On Error GoTo ErrorHandler
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 1
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleSTFndPort1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleSTFndPort1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleSTFndPort1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleSTFndPort1"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleSTFndPort1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleSTFndPort1"
    On Error GoTo ErrorHandler

    Dim objEqpFoundationPort As IJEqpFoundationPort
    Dim holes() As Variant
    
    'Get the Symbol inputs of the Component
    GetDimensionsFromSymbolArray oPropertyDescription.CAO
    
    'Now Construct the holes Array
    Set objEqpFoundationPort = oPropertyDescription.Object
    Call objEqpFoundationPort.GetHoles(holes())
          
    holes(0, 1) = -m_dFP1XBoltHole
    holes(0, 2) = -m_dFP1YBoltHole
    holes(1, 1) = m_dFP1XBoltHole
    holes(1, 2) = -m_dFP1YBoltHole
    holes(2, 1) = m_dFP1XBoltHole
    holes(2, 2) = m_dFP1YBoltHole
    holes(3, 1) = -m_dFP1XBoltHole
    holes(3, 2) = m_dFP1YBoltHole

    'Set the calculated holes on foundation port
    Call objEqpFoundationPort.SetHoles(holes)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleSTFndPort1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleSTFndPort1"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 1
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleSTFndPort1(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleSTFndPort1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleSTFndPort1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleSTFndPort1"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleSTFndPort2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleSTFndPort2(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleSTFndPort2"
    On Error GoTo ErrorHandler
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 2
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleSTFndPort2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleSTFndPort2"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleSTFndPort2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleSTFndPort2"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleSTFndPort2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleSTFndPort2"
    On Error GoTo ErrorHandler
    
    Dim objEqpFoundationPort As IJEqpFoundationPort
    Dim holes() As Variant
    
    'Get the Symbol inputs of the Component
    GetDimensionsFromSymbolArray oPropertyDescription.CAO
    
    'Now Construct the holes Array
    Set objEqpFoundationPort = oPropertyDescription.Object
    Call objEqpFoundationPort.GetHoles(holes())
    
    holes(0, 1) = -m_dFP2XBoltHole
    holes(0, 2) = -m_dFP2YBoltHole
    holes(1, 1) = m_dFP2XBoltHole
    holes(1, 2) = -m_dFP2YBoltHole
    holes(2, 1) = m_dFP2XBoltHole
    holes(2, 2) = m_dFP2YBoltHole
    holes(3, 1) = -m_dFP2XBoltHole
    holes(3, 2) = m_dFP2YBoltHole
 
    'Set the calculated holes on foundation port
    Call objEqpFoundationPort.SetHoles(holes)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleSTFndPort2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleSTFndPort2"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 2
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleSTFndPort2(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleSTFndPort2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleSTFndPort2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleSTFndPort2"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for Variable Front Ends
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructTyACDN(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructTyACDN"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO
        
    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "FrontEndTypeACDN 01-EC", "TyACDN")
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oPosition As IJDPosition
    Set oPosition = New AutoMath.DPosition
    oPosition.Set 0, 0, 0
    
    'To set the End according to the Occurnece Attributes
    TransformEnd oDesignEquipment, pObject, oPosition
    
    Set oPosition = Nothing
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
  
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTyACDN(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructTyACDN"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsTyACDN(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsTyACDN"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeDia").Value = m_dFrontEndFlangeDia
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeTk1").Value = m_dFrontEndFlangeTk1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength1").Value = m_dFrontEndLength1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength2").Value = m_dFrontEndLength2
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeTk2").Value = m_dFrontEndFlangeTk2
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeTk3").Value = m_dFrontEndFlangeTk3
    oAttribs.CollectionOfAttributes("IJUAEXCHANGERCHANNEL").Item("ChannelDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
       
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateTyACDN(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateTyACDN"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryTyACDN(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryTyACDN"
    On Error GoTo ErrorHandler
    LogCalls METHOD
       
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalTyACDN(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalTyACDN"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dFrontEndType = FrontEndType_A Or m_dFrontEndType = FrontEndType_C Or _
        m_dFrontEndType = FrontEndType_D Or m_dFrontEndType = FrontEndType_N Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseTyACDN(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseTyACDN"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConstructTyB(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructTyB"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO
     
    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "FrontEndTypeB 01-EC", "TyB")
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oPosition As IJDPosition
    Set oPosition = New AutoMath.DPosition
    oPosition.Set 0, 0, 0
    
    'To set the End according to the Occurnece Attributes
    TransformEnd oDesignEquipment, pObject, oPosition
    
    Set oPosition = Nothing
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTyB(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructTyB"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsTyB(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsTyB"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeDia").Value = m_dFrontEndFlangeDia
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeTk1").Value = m_dFrontEndFlangeTk1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength1").Value = m_dFrontEndLength1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength2").Value = m_dFrontEndLength2
    oAttribs.CollectionOfAttributes("IJUAEXCHANGERCHANNEL").Item("ChannelDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateTyB(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateTyB"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryTyB(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryTyB"
    On Error GoTo ErrorHandler
    LogCalls METHOD
        
  Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalTyB(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalTyB"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dFrontEndType = FrontEndType_B Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseTyB(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseTyB"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConstructTyFrontQ(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructTyFrontQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO
    
    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "FrontEndTypeQ 01-EC", "TyQ")
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oPosition As IJDPosition
    Set oPosition = New AutoMath.DPosition
    oPosition.Set 0, 0, 0
    
    'To set the End according to the Occurnece Attributes
    TransformEnd oDesignEquipment, pObject, oPosition
    
    Set oPosition = Nothing
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTyFrontQ(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructTyFrontQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsTyFrontQ(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsTyFrontQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject
    
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeDia").Value = m_dFrontEndFlangeDia
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeTk1").Value = m_dFrontEndFlangeTk1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength1").Value = m_dFrontEndLength1
    oAttribs.CollectionOfAttributes("IJUAEXCHANGERCHANNEL").Item("FrontEndDiameter").Value = m_dFrontEndDiameter
    oAttribs.CollectionOfAttributes("IJUAEXCHANGERCHANNEL").Item("ChannelDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    
    m_bComputeEqpComp = False
 
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateTyFrontQ(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateTyFrontQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryTyFrontQ(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryTyFrontQ"
    On Error GoTo ErrorHandler
    LogCalls METHOD
         
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalTyFrontQ(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalTyFrontQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dFrontEndType = FrontEndType_Q Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseTyFrontQ(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseTyFrontQ"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For EndType2T01
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEndType2T01(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEndType2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    Dim oPosition As IJDPosition
    
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndType2TO1-EC", _
                             "End Type 2TO1")
                             
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEndType2T01(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEndType2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEndType2T01(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEndType2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject
    
    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEndType2T01(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEndType2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEndType2T01(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEndType2T01"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEqpCompEquip = oEqpComponent

    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oEquipment = Nothing
    
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEndType2T01(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEndType2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dVesselHeadType = EndType_2TO1 Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEndType2T01(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEndType2T01"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For EndTypeCone
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEndTypeCone(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEndTypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    
    Set oDesignEquipment = pMemberDescription.CAO
    
     

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeCone-EC", _
                             "End Type Cone")
                             
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oDesignEquipment = Nothing
    
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEndTypeCone(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEndTypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEndTypeCone(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEndTypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeTopDiameter").Value = m_dEndHeadConeTopDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeHeight").Value = m_dEndHeadConeHeight
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEndTypeCone(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEndTypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEndTypeCone(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEndTypeCone"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEqpCompEquip = oEqpComponent
    
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oEquipment = Nothing
    
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEndTypeCone(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEndTypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dVesselHeadType = EndType_CONE Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEndTypeCone(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEndTypeCone"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For EndTypeDone
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEndTypeDome(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEndTypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    
    Set oDesignEquipment = pMemberDescription.CAO
    
    

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeDome-EC", _
                                                    "End Type Dome")
                                                    
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oDesignEquipment = Nothing
    
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEndTypeDome(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEndTypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEndTypeDome(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEndTypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject
    
    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadDomeradius").Value = m_dEndHeadDomeradius
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    
    m_bComputeEqpComp = False
 
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEndTypeDome(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEndTypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEndTypeDome(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEndTypeDome"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEqpCompEquip = oEqpComponent
    
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oEquipment = Nothing

    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEndTypeDome(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEndTypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dVesselHeadType = EndType_DOME Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEndTypeDome(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEndTypeDome"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For EndTypeFND
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEndTypeFND(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEndTypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    
    Set oDesignEquipment = pMemberDescription.CAO
    
    

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeFnD-EC", _
                                    "End Type F&D")
                                    
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oDesignEquipment = Nothing
    
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEndTypeFND(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEndTypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEndTypeFND(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEndTypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEndTypeFND(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEndTypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEndTypeFND(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEndTypeFND"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEqpCompEquip = oEqpComponent
    
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oEquipment = Nothing
    
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEndTypeFND(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEndTypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dVesselHeadType = EndType_FnD Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEndTypeFND(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEndTypeFND"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For EndTypeFLGD
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEndTypeFLGD(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEndTypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    
    Set oDesignEquipment = pMemberDescription.CAO
    
    

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeFLGD-EC", _
                                    "End Type FLGD")
                                    
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oDesignEquipment = Nothing
    
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEndTypeFLGD(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEndTypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEndTypeFLGD(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEndTypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject
    
    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadFlangedThick1").Value = m_dEndHeadFlangedThick1
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadFlangedThick2").Value = m_dEndHeadFlangedThick2
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
       
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEndTypeFLGD(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEndTypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEndTypeFLGD(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEndTypeFLGD"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEqpCompEquip = oEqpComponent
    
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oEquipment = Nothing
    
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEndTypeFLGD(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEndTypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dVesselHeadType = EndType_FLGD Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEndTypeFLGD(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEndTypeFLGD"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For EndTypeHemi
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEndTypeHemi(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEndTypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    
    Set oDesignEquipment = pMemberDescription.CAO
    
    

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeHemi-EC", _
                                    "End Type Hemi")
                                    
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oDesignEquipment = Nothing
    
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEndTypeHemi(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEndTypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEndTypeHemi(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEndTypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
  
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEndTypeHemi(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEndTypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEndTypeHemi(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEndTypeHemi"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEqpCompEquip = oEqpComponent
    
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oEquipment = Nothing
    
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEndTypeHemi(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEndTypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dVesselHeadType = EndType_HEMI Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEndTypeHemi(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEndTypeHemi"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For EndTypeTORC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEndTypeTORC(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEndTypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    
    Set oDesignEquipment = pMemberDescription.CAO
    
    

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeTORC-EC", _
                                    "End Type TORC")
                                    
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, 0, -PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oDesignEquipment = Nothing
    
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEndTypeTORC(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEndTypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEndTypeTORC(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEndTypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeTopDiameter").Value = m_dEndHeadConeTopDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeHeight").Value = m_dEndHeadConeHeight
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadKnuckleRadius").Value = m_dEndHeadKnuckleRadius
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEndTypeTORC(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEndTypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEndTypeTORC(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEndTypeTORC"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEqpCompEquip = oEqpComponent
    
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpComponent, 0, -PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oEquipment = Nothing
    
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEndTypeTORC(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEndTypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dVesselHeadType = EndType_TORC Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEndTypeTORC(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEndTypeTORC"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For EndTypeTORS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEndTypeTORS(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEndTypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    
    Set oDesignEquipment = pMemberDescription.CAO
    
    

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeTORS-EC", _
                                    "End Type TORS")
                                    
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, 0, -PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oDesignEquipment = Nothing
    
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEndTypeTORS(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEndTypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEndTypeTORS(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEndTypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadSphericalRadius").Value = m_dEndHeadSphericalRadius
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadKnuckleRadius").Value = m_dEndHeadKnuckleRadius
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEndTypeTORS(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEndTypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEndTypeTORS(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEndTypeTORS"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEqpCompEquip = oEqpComponent
    
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpComponent, 0, -PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oEquipment = Nothing
    
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEndTypeTORS(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEndTypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dVesselHeadType = EndType_TORS Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEndTypeTORS(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEndTypeTORS"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For EndTypeFLAT
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEndTypeFLAT(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEndTypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeFLAT-EC", _
                                    "End Type FLAT")
                                    
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oDesignEquipment = Nothing
    
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEndTypeFLAT(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEndTypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEndTypeFLAT(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEndTypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEndTypeFLAT(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEndTypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEndTypeFLAT(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEndTypeFLAT"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEqpCompEquip = oEqpComponent
    
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dFrontEndLength1 + m_dExchangerLength, 0, (m_dExchangerDiameter / 2 - m_dExchangerNeckDiameter / 2)
    
    Set oEquipment = Nothing
    
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEndTypeFLAT(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEndTypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dVesselHeadType = EndType_FLAT Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEndTypeFLAT(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEndTypeFLAT"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub TransformAndRotate(Equipment As IJEquipment, Object As Object, _
                                    dAngleOfRotAboutX As Double, _
                                    dAngleOfRotAboutY As Double, dAngleOfRotAboutZ As Double, _
                                    dXOffset As Double, dYOffset As Double, dZOffset As Double)
    Const METHOD = "TransformAndRotate"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    If Not Object Is Nothing Then
        If TypeOf Object Is IJDGeometry Then
            Dim oEqpMatrix As IJDT4x4
            Dim oTransform As IJEquipment
            Dim oDirVector As New AutoMath.DVector
            Dim oEquipPrimary As New AutoMath.DVector
            Dim oEquipSecondary As New AutoMath.DVector
            Dim oEquipNormal As New AutoMath.DVector
            
            Equipment.GetMatrix oEqpMatrix
            oEquipPrimary.Set oEqpMatrix.IndexValue(0), oEqpMatrix.IndexValue(1), oEqpMatrix.IndexValue(2)
            oEquipSecondary.Set oEqpMatrix.IndexValue(4), oEqpMatrix.IndexValue(5), oEqpMatrix.IndexValue(6)
            oEquipNormal.Set oEqpMatrix.IndexValue(8), oEqpMatrix.IndexValue(9), oEqpMatrix.IndexValue(10)
            

            Set oTransform = Object
            If Not dAngleOfRotAboutX = 0 Then
                oDirVector.Set 1, 0, 0
                oEqpMatrix.Rotate dAngleOfRotAboutX, oDirVector
            End If
            If Not dAngleOfRotAboutY = 0 Then
                oDirVector.Set 0, 1, 0
                oEqpMatrix.Rotate dAngleOfRotAboutY, oDirVector
            End If
            If Not dAngleOfRotAboutZ = 0 Then
                oDirVector.Set 0, 0, 1
                oEqpMatrix.Rotate dAngleOfRotAboutZ, oDirVector
            End If
            'oTransform.DTransform oEqpMatrix
            oTransform.SetMatrix oEqpMatrix
            
            
            
            Dim oCompOrigin As IJDPosition
            Dim x As Double, y As Double, z As Double
            
            Set oCompOrigin = New AutoMath.DPosition
            Equipment.GetPosition x, y, z
            oCompOrigin.Set x, y, z
            
            If dXOffset <> 0# Then
                oEquipPrimary.Length = dXOffset
                Set oCompOrigin = oCompOrigin.Offset(oEquipPrimary)
            End If
            
            'oOrigin.Set oCompOrigin.x, oCompOrigin.y, oCompOrigin.z
            If dYOffset <> 0# Then
                oEquipSecondary.Length = dYOffset
                Set oCompOrigin = oCompOrigin.Offset(oEquipSecondary)
            End If
            'oOrigin.Set oCompOrigin.x, oCompOrigin.y, oCompOrigin.z
            If dZOffset <> 0# Then
                oEquipNormal.Length = dZOffset
                Set oCompOrigin = oCompOrigin.Offset(oEquipNormal)
            End If
            
            oTransform.SetPosition oCompOrigin.x, oCompOrigin.y, oCompOrigin.z
            
            Set oTransform = Nothing
            Set oEqpMatrix = Nothing
            Set oDirVector = Nothing
        End If
    End If
     
    Exit Sub
ErrorHandler:
    Set oTransform = Nothing
    Set oEqpMatrix = Nothing
    Set oDirVector = Nothing
    HandleError MODULE, METHOD
End Sub

Private Sub TransformEnd(Equipment As IJEquipment, Object As Object, oPosition As IJDPosition)
    Const METHOD = "TransformEnd"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oEqpMatrix As IJDT4x4
    Dim oTransform As IJDGeometry
    If Not Object Is Nothing Then
        If TypeOf Object Is IJDGeometry Then
                Equipment.GetMatrix oEqpMatrix
                If Not oPosition Is Nothing Then
                    oEqpMatrix.IndexValue(12) = oEqpMatrix.IndexValue(12) + oPosition.x
                    oEqpMatrix.IndexValue(13) = oEqpMatrix.IndexValue(13) + oPosition.y
                    oEqpMatrix.IndexValue(14) = oEqpMatrix.IndexValue(14) + oPosition.z
                End If
                Set oTransform = Object
                oTransform.DTransform oEqpMatrix
       End If
    End If

    Set oTransform = Nothing
    Set oEqpMatrix = Nothing
    
    Exit Sub
ErrorHandler:
    Set oTransform = Nothing
    Set oEqpMatrix = Nothing
    HandleError MODULE, METHOD
End Sub



























