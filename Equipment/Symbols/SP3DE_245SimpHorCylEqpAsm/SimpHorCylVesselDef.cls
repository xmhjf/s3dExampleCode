VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HorCylDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"CAD"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:28:54 PM"
Attribute VB_Ext_KEY = "SP3DEqpCADTemplateVersion" ,"1.2"
'******************************************************************************
'   Copyright (C) 2007, Intergraph Corporation. All rights reserved.
'
'   File:      SimpHorCylVesselDef.cls
'   Author:    RUK
'   Creation Date:  Wednesday, March 7 2007
'
'   Description:
'   This is E245 Simple Horizontal Cylindrical Equipment Assembly.
'   This Assembly has two DatumPoints with Orientations X towards Up and Y along North, and
'   One Circular FoundationPort
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "SP3DE_245SimpHorCylEqpAsm:HorCylDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

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

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors

Private m_avSymbolArrayOfInputs()   As Variant

Private m_oNorth                    As IJDVector
Private m_oEast                     As IJDVector
Private m_oElevation                As IJDVector

Private m_dEnd1Type As Long
Private m_dEnd2Type As Long
Private m_dVesselTantoTan As Double
Private m_dVesselDiameter As Double
Private m_dVesselCenterHeight As Double
Private m_dFirstSupportLocation As Double
Private m_dSecondSupportLocation As Double
Private m_dXBoltHole As Double
Private m_dYBoltHole As Double
Private m_dEndHeadConeTopDiameter As Double
Private m_dEndHeadConeHeight As Double
Private m_dEndHeadKnuckleRadius As Double
Private m_dEndHeadDomeradius As Double
Private m_dEndHeadFlangedThick1 As Double
Private m_dEndHeadFlangedThick2 As Double
Private m_dEndHeadSphericalRadius As Double
Private m_dInsulationThickness As Double

Private PI As Double
Private m_bComputeEqpComp As Boolean

' Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt

Private Sub Class_Initialize()
    Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
    
    m_oEquipCADHelper.ProjectName = "SP3DE_245SimpHorCylEqpAsm"
    m_oEquipCADHelper.ClassName = "HorCylDef"
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
        
    'Add new member(NozzleSTFndPort1) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleSTFndPort1", 5, "CMConstructNozzleSTFndPort1", imsCOOKIE_ID_USS_LIB)
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
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleSTFndPort2", 6, "CMConstructNozzleSTFndPort2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleSTFndPort2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleSTFndPort2"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleSTFndPort2"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleSTFndPort2"
    
    'Add properties for (NozzleSTFndPort2)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleSTFndPort2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleSTFndPort2", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleSTFndPort2GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleSTFndPort2", imsCOOKIE_ID_USS_LIB
    
    'Vessel Heads
    'Add new member(End1Type2T01) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End1Type2T01", 7, "CMConstructEnd1Type2T01", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd1Type2T01"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd1Type2T01"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd1Type2T01"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd1Type2T01"
    'Add properties for (End1Type2T01)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End1Type2T01Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd1Type2T01", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End1Type2T01GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd1Type2T01", imsCOOKIE_ID_USS_LIB
    
    'Add new member(End1TypeCone) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End1TypeCone", 8, "CMConstructEnd1TypeCone", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd1TypeCone"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd1TypeCone"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd1TypeCone"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd1TypeCone"
    'Add properties for (End1TypeCone)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End1TypeConeProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd1TypeCone", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End1TypeConeGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd1TypeCone", imsCOOKIE_ID_USS_LIB

    'Add new member(End1TypeDome) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End1TypeDome", 9, "CMConstructEnd1TypeDome", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd1TypeDome"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd1TypeDome"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd1TypeDome"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd1TypeDome"
    'Add properties for (End1TypeDome)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End1TypeDomeProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd1TypeDome", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End1TypeDomeGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd1TypeDome", imsCOOKIE_ID_USS_LIB

    'Add new member(End1TypeFND) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End1TypeFND", 10, "CMConstructEnd1TypeFND", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd1TypeFND"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd1TypeFND"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd1TypeFND"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd1TypeFND"
    'Add properties for (End1TypeFND)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End1TypeFNDProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd1TypeFND", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End1TypeFNDGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd1TypeFND", imsCOOKIE_ID_USS_LIB
    
    'Add new member(End1TypeFLGD) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End1TypeFLGD", 11, "CMConstructEnd1TypeFLGD", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd1TypeFLGD"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd1TypeFLGD"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd1TypeFLGD"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd1TypeFLGD"
    'Add properties for (End1TypeFLGD)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End1TypeFLGDProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd1TypeFLGD", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End1TypeFLGDGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd1TypeFLGD", imsCOOKIE_ID_USS_LIB

    'Add new member(End1TypeHemi) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End1TypeHemi", 12, "CMConstructEnd1TypeHemi", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd1TypeHemi"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd1TypeHemi"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd1TypeHemi"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd1TypeHemi"
    'Add properties for (End1TypeHemi)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End1TypeHemiProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd1TypeHemi", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End1TypeHemiGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd1TypeHemi", imsCOOKIE_ID_USS_LIB

    'Add new member(End1TypeTORC) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End1TypeTORC", 13, "CMConstructEnd1TypeTORC", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd1TypeTORC"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd1TypeTORC"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd1TypeTORC"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd1TypeTORC"
    'Add properties for (End1TypeTORC)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End1TypeTORCProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd1TypeTORC", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End1TypeTORCGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd1TypeTORC", imsCOOKIE_ID_USS_LIB

    'Add new member(End1TypeTORS) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End1TypeTORS", 14, "CMConstructEnd1TypeTORS", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd1TypeTORS"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd1TypeTORS"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd1TypeTORS"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd1TypeTORS"
    'Add properties for (End1TypeTORS)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End1TypeTORSProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd1TypeTORS", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End1TypeTORSGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd1TypeTORS", imsCOOKIE_ID_USS_LIB

    'Add new member(End1TypeFLAT) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End1TypeFLAT", 15, "CMConstructEnd1TypeFLAT", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd1TypeFLAT"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd1TypeFLAT"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd1TypeFLAT"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd1TypeFLAT"
    'Add properties for (End1TypeFLAT)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End1TypeFLATProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd1TypeFLAT", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End1TypeFLATGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd1TypeFLAT", imsCOOKIE_ID_USS_LIB

    'Add new member(End2Type2T01) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End2Type2T01", 16, "CMConstructEnd2Type2T01", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd2Type2T01"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd2Type2T01"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd2Type2T01"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd2Type2T01"
    'Add properties for (End2Type2T01)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End2Type2T01Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd2Type2T01", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End2Type2T01GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd2Type2T01", imsCOOKIE_ID_USS_LIB

    'Add new member(End2TypeCone) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End2TypeCone", 17, "CMConstructEnd2TypeCone", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd2TypeCone"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd2TypeCone"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd2TypeCone"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd2TypeCone"
    'Add properties for (End2TypeCone)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End2TypeConeProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd2TypeCone", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End2TypeConeGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd2TypeCone", imsCOOKIE_ID_USS_LIB

    'Add new member(End2TypeDome) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End2TypeDome", 18, "CMConstructEnd2TypeDome", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd2TypeDome"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd2TypeDome"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd2TypeDome"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd2TypeDome"
    'Add properties for (End2TypeDome)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End2TypeDomeProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd2TypeDome", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End2TypeDomeGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd2TypeDome", imsCOOKIE_ID_USS_LIB

    'Add new member(End2TypeFND) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End2TypeFND", 19, "CMConstructEnd2TypeFND", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd2TypeFND"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd2TypeFND"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd2TypeFND"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd2TypeFND"
    'Add properties for (End2TypeFND)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End2TypeFNDProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd2TypeFND", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End2TypeFNDGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd2TypeFND", imsCOOKIE_ID_USS_LIB

    'Add new member(End2TypeFLGD) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End2TypeFLGD", 20, "CMConstructEnd2TypeFLGD", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd2TypeFLGD"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd2TypeFLGD"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd2TypeFLGD"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd2TypeFLGD"
    'Add properties for (End2TypeFLGD)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End2TypeFLGDProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd2TypeFLGD", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End2TypeFLGDGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd2TypeFLGD", imsCOOKIE_ID_USS_LIB

   'Add new member(End2TypeHemi) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End2TypeHemi", 21, "CMConstructEnd2TypeHemi", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd2TypeHemi"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd2TypeHemi"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd2TypeHemi"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd2TypeHemi"
    'Add properties for (End2TypeHemi)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End2TypeHemiProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd2TypeHemi", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End2TypeHemiGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd2TypeHemi", imsCOOKIE_ID_USS_LIB

    'Add new member(End2TypeTORC) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End2TypeTORC", 22, "CMConstructEnd2TypeTORC", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd2TypeTORC"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd2TypeTORC"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd2TypeTORC"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd2TypeTORC"
    'Add properties for (End2TypeTORC)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End2TypeTORCProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd2TypeTORC", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End2TypeTORCGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd2TypeTORC", imsCOOKIE_ID_USS_LIB

    'Add new member(End2TypeTORS) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End2TypeTORS", 23, "CMConstructEnd2TypeTORS", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd2TypeTORS"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd2TypeTORS"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd2TypeTORS"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd2TypeTORS"
    'Add properties for (End2TypeTORS)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End2TypeTORSProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd2TypeTORS", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End2TypeTORSGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd2TypeTORS", imsCOOKIE_ID_USS_LIB

    'Add new member(End2TypeFLAT) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("End2TypeFLAT", 24, "CMConstructEnd2TypeFLAT", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEnd2TypeFLAT"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEnd2TypeFLAT"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEnd2TypeFLAT"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEnd2TypeFLAT"
    'Add properties for (End2TypeFLAT)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "End2TypeFLATProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEnd2TypeFLAT", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "End2TypeFLATGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryEnd2TypeFLAT", imsCOOKIE_ID_USS_LIB

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
        Case "DP1", "DP2", "DP3", "DP4", "NozzleSTFndPort1", "NozzleSTFndPort2", _
                "End1Type2T01", "End1TypeCone", "End1TypeDome", "End1TypeFND", "End1TypeFLGD", "End1TypeHemi", "End1TypeTORC", "End1TypeTORS", "End1TypeFLAT", _
                "End2Type2T01", "End2TypeCone", "End2TypeDome", "End2TypeFND", "End2TypeFLGD", "End2TypeHemi", "End2TypeTORC", "End2TypeTORS", "End2TypeFLAT"
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
        Case "DP1", "DP2", "DP3", "DP4"
            For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                    Case "IJUADATUMSHAPE"
                        oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                    Case Else
                        '
                End Select
            Next
        Case "End1Type2T01", "End1TypeCone", "End1TypeDome", "End1TypeFND", "End1TypeFLGD", "End1TypeHemi", "End1TypeTORC", "End1TypeTORS", "End1TypeFLAT", _
            "End2Type2T01", "End2TypeCone", "End2TypeDome", "End2TypeFND", "End2TypeFLGD", "End2TypeHemi", "End2TypeTORC", "End2TypeTORS", "End2TypeFLAT"
                For m = 1 To oAttrCollection.Count
                    Set oAttributeDescriptor = oAttrCollection.Item(m)
                    Select Case UCase(oAttributeDescriptor.InterfaceName)
                        Case "IJUAEquipmentEnd", "IJUAVesselDiameter", "IJInsulationThickness"
                            oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsChanged
                        Case Else
                            '
                    End Select
                Next
        Case Else
            '
    End Select
    Select Case oMemberDescription.Name
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

Public Sub CMEvaluateEquipment(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEquipment"
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
          
    holes(0, 1) = -m_dXBoltHole
    holes(0, 2) = -m_dYBoltHole
    holes(1, 1) = m_dXBoltHole
    holes(1, 2) = -m_dYBoltHole
    holes(2, 1) = m_dXBoltHole
    holes(2, 2) = m_dYBoltHole
    holes(3, 1) = -m_dXBoltHole
    holes(3, 2) = m_dYBoltHole

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
     
    holes(0, 1) = -m_dXBoltHole
    holes(0, 2) = -m_dYBoltHole
    holes(1, 1) = m_dXBoltHole
    holes(1, 2) = -m_dYBoltHole
    holes(2, 1) = m_dXBoltHole
    holes(2, 2) = m_dYBoltHole
    holes(3, 1) = -m_dXBoltHole
    holes(3, 2) = m_dYBoltHole
    
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

'The code of this method is specific to each Equipment as it converts the array of inputs in a set of global variables
Private Sub GetDimensionsFromSymbolArray(SmartOccurrence As IJSmartOccurrence)
    Const METHOD = "GetDimensionsFromSymbolArray"
    On Error GoTo ErrorHandler
    
    m_avSymbolArrayOfInputs = m_oEquipCADHelper.GetSymbolArrayOfInputs(SmartOccurrence)

    'Inputs,  from equipment symbol code
    m_dEnd1Type = m_avSymbolArrayOfInputs(2)
    m_dEnd2Type = m_avSymbolArrayOfInputs(3)
    m_dVesselTantoTan = m_avSymbolArrayOfInputs(4)          'P2
    m_dVesselDiameter = m_avSymbolArrayOfInputs(5)          'P3
    m_dVesselCenterHeight = m_avSymbolArrayOfInputs(6)      'P4
    m_dFirstSupportLocation = m_avSymbolArrayOfInputs(7)    'P5
    m_dSecondSupportLocation = m_avSymbolArrayOfInputs(8)   'P6
    m_dXBoltHole = m_avSymbolArrayOfInputs(11)
    m_dYBoltHole = m_avSymbolArrayOfInputs(12)
    m_dEndHeadConeTopDiameter = m_avSymbolArrayOfInputs(13)
    m_dEndHeadConeHeight = m_avSymbolArrayOfInputs(14)
    m_dEndHeadKnuckleRadius = m_avSymbolArrayOfInputs(15)
    m_dEndHeadDomeradius = m_avSymbolArrayOfInputs(16)
    m_dEndHeadFlangedThick1 = m_avSymbolArrayOfInputs(17)
    m_dEndHeadFlangedThick2 = m_avSymbolArrayOfInputs(18)
    m_dEndHeadSphericalRadius = m_avSymbolArrayOfInputs(19)
    m_dInsulationThickness = m_avSymbolArrayOfInputs(20)

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
    'The X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'The Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oElevation
    
    Set oPosition = Nothing
End Sub

Private Sub PositionAndOrientDP2(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    oPosition.Set m_dVesselTantoTan, 0, 0
    
    'We want:
    'The X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'The Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oElevation
    
    Set oPosition = Nothing
End Sub

Private Sub PositionAndOrientDP3(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    oPosition.Set m_dFirstSupportLocation, 0, -m_dVesselCenterHeight
    
    'We want:
    'The X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'The Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oElevation
    
    Set oPosition = Nothing
End Sub

Private Sub PositionAndOrientDP4(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    oPosition.Set m_dFirstSupportLocation + m_dSecondSupportLocation, 0, -m_dVesselCenterHeight
    
    'We want:
    'The X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'The Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oElevation
    
    Set oPosition = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Vessel  Heads
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End1Type2T01
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd1Type2T01(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd1Type2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndType2TO1-EC", _
                                    "End Type 2TO1")
    'Transform the Vessel Head to DP1 and Rotate to 90 degrees c.c.w direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2, -PI / 2, 0, 0, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd1Type2T01(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd1Type2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd1Type2T01(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd1Type2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd1Type2T01(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd1Type2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd1Type2T01(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd1Type2T01"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpCompEquip, PI / 2#, -PI / 2#, 0, 0, 0, 0

    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd1Type2T01(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd1Type2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd1Type = EndType_2TO1 Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd1Type2T01(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd1Type2T01"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End1TypeCone
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd1TypeCone(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd1TypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeCone-EC", _
                             "End Type Cone")
    'Transform the Vessel Head to DP1 and Rotate to 90 degrees c.c.w direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2, -PI / 2, 0, 0, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd1TypeCone(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd1TypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd1TypeCone(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd1TypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeTopDiameter").Value = m_dEndHeadConeTopDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeHeight").Value = m_dEndHeadConeHeight
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd1TypeCone(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd1TypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd1TypeCone(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd1TypeCone"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpCompEquip, PI / 2#, -PI / 2#, 0, 0, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
   
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd1TypeCone(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd1TypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd1Type = EndType_CONE Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd1TypeCone(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd1TypeCone"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End1TypeDome
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd1TypeDome(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd1TypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeDome-EC", _
                                                    "End Type Dome")
    'Transform the Vessel Head to DP1 and Rotate to 90 degrees c.c.w direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2, -PI / 2, 0, 0, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd1TypeDome(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd1TypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd1TypeDome(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd1TypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject
    
    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadDomeradius").Value = m_dEndHeadDomeradius
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd1TypeDome(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd1TypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd1TypeDome(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd1TypeDome"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpCompEquip, PI / 2#, -PI / 2#, 0, 0, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
   
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd1TypeDome(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd1TypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd1Type = EndType_DOME Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Public Sub CMReleaseEnd1TypeDome(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd1TypeDome"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End1TypeFND
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd1TypeFND(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd1TypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeFnD-EC", _
                                    "End Type FnD")
    'Transform the Vessel Head to DP1 and Rotate to 90 degrees c.c.w direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2, -PI / 2, 0, 0, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd1TypeFND(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd1TypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd1TypeFND(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd1TypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd1TypeFND(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd1TypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd1TypeFND(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd1TypeFND"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpCompEquip, PI / 2#, -PI / 2#, 0, 0, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
   
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd1TypeFND(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd1TypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd1Type = EndType_FnD Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd1TypeFND(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd1TypeFND"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End1TypeFLGD
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd1TypeFLGD(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd1TypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeFLGD-EC", _
                                    "End Type FLGD")
    'Transform the Vessel Head to DP1 and Rotate to 90 degrees c.c.w direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2, -PI / 2, 0, 0, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd1TypeFLGD(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd1TypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd1TypeFLGD(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd1TypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadFlangedThick1").Value = m_dEndHeadFlangedThick1
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadFlangedThick2").Value = m_dEndHeadFlangedThick2
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
       
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd1TypeFLGD(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd1TypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd1TypeFLGD(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd1TypeFLGD"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpCompEquip, PI / 2#, -PI / 2#, 0, 0, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
   
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd1TypeFLGD(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd1TypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd1Type = EndType_FLGD Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd1TypeFLGD(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd1TypeFLGD"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End1TypeHemi
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd1TypeHemi(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd1TypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeHemi-EC", _
                                    "End Type Hemi")
    'Transform the Vessel Head to DP1 and Rotate to 90 degrees c.c.w direction w.r.t North
    'TransformAndRotate oDesignEquipment, pObject, oPosition, PI / 2#, -PI / 2#, 0
    TransformAndRotate oDesignEquipment, pObject, PI / 2, -PI / 2, 0, 0, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd1TypeHemi(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd1TypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd1TypeHemi(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd1TypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd1TypeHemi(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd1TypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd1TypeHemi(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd1TypeHemi"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpCompEquip, PI / 2#, -PI / 2#, 0, 0, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
   
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd1TypeHemi(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd1TypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd1Type = EndType_HEMI Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd1TypeHemi(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd1TypeHemi"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End1TypeTORC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd1TypeTORC(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd1TypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeTORC-EC", _
                                    "End Type TORC")
    'Transform the Vessel Head to DP1 and Rotate to 90 degrees c.c.w direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, 0, PI / 2, 0, 0, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd1TypeTORC(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd1TypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd1TypeTORC(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd1TypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeTopDiameter").Value = m_dEndHeadConeTopDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeHeight").Value = m_dEndHeadConeHeight
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadKnuckleRadius").Value = m_dEndHeadKnuckleRadius
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd1TypeTORC(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd1TypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd1TypeTORC(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd1TypeTORC"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpCompEquip, 0, PI / 2#, 0, 0, 0, 0

    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
   
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd1TypeTORC(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd1TypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd1Type = EndType_TORC Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd1TypeTORC(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd1TypeTORC"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End1TypeTORS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd1TypeTORS(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd1TypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeTORS-EC", _
                                    "End Type TORS")
    'Transform the Vessel Head to DP1 and Rotate to 90 degrees c.c.w direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, 0, PI / 2, 0, 0, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd1TypeTORS(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd1TypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd1TypeTORS(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd1TypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadSphericalRadius").Value = m_dEndHeadSphericalRadius
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadKnuckleRadius").Value = m_dEndHeadKnuckleRadius
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd1TypeTORS(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd1TypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd1TypeTORS(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd1TypeTORS"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpCompEquip, 0, PI / 2#, 0, 0, 0, 0

    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
   
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd1TypeTORS(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd1TypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd1Type = EndType_TORS Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd1TypeTORS(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd1TypeTORS"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End1TypeFLAT
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd1TypeFLAT(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd1TypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeFLAT-EC", _
                                    "End Type FLAT")
    'Transform the Vessel Head to DP1 and Rotate to 90 degrees c.c.w direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2, -PI / 2, 0, 0, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd1TypeFLAT(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd1TypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd1TypeFLAT(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd1TypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd1TypeFLAT(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd1TypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd1TypeFLAT(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd1TypeFLAT"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oEquipment As IJEquipment
    Dim oEqpCompEquip As IJEquipment

    Set oEqpComponent = oPropertyDescription.Object
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    Set oEqpCompEquip = oEqpComponent
    oEqpCompEquip.SetPosition 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformAndRotate oEquipment, oEqpCompEquip, PI / 2#, -PI / 2#, 0, 0, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd1TypeFLAT(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd1TypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd1Type = EndType_FLAT Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd1TypeFLAT(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd1TypeFLAT"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End2Type2T01
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd2Type2T01(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd2Type2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    Set oDesignEquipment = pMemberDescription.CAO
     
    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndType2TO1-EC", _
                             "End Type 2TO1")
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2, PI / 2, 0, m_dVesselTantoTan, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd2Type2T01(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd2Type2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd2Type2T01(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd2Type2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject
    
    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd2Type2T01(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd2Type2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd2Type2T01(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd2Type2T01"
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
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd2Type2T01(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd2Type2T01"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd2Type = EndType_2TO1 Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd2Type2T01(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd2Type2T01"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End2TypeCone
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd2TypeCone(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd2TypeCone"
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
    TransformAndRotate oDesignEquipment, pObject, PI / 2, PI / 2, 0, m_dVesselTantoTan, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd2TypeCone(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd2TypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd2TypeCone(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd2TypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeTopDiameter").Value = m_dEndHeadConeTopDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeHeight").Value = m_dEndHeadConeHeight
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd2TypeCone(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd2TypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd2TypeCone(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd2TypeCone"
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
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd2TypeCone(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd2TypeCone"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd2Type = EndType_CONE Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd2TypeCone(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd2TypeCone"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End2TypeDone
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd2TypeDome(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd2TypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    Set oDesignEquipment = pMemberDescription.CAO
     
    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeDome-EC", _
                                                    "End Type Dome")
    'TransformAndRotate oDesignEquipment, pObject, oPosition, PI / 2#, PI / 2#, 0
    TransformAndRotate oDesignEquipment, pObject, PI / 2, PI / 2, 0, m_dVesselTantoTan, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd2TypeDome(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd2TypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd2TypeDome(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd2TypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject
    
    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadDomeradius").Value = m_dEndHeadDomeradius
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    
    m_bComputeEqpComp = False
 
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd2TypeDome(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd2TypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd2TypeDome(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd2TypeDome"
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
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd2TypeDome(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd2TypeDome"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd2Type = EndType_DOME Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd2TypeDome(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd2TypeDome"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End2TypeFND
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd2TypeFND(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd2TypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    
    Set oDesignEquipment = pMemberDescription.CAO
     
    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "EndTypeFnD-EC", _
                                    "End Type FnD")
    'Transform the Head to DP2 and Rotate to 90 degrees in clockwise direction w.r.t North
    TransformAndRotate oDesignEquipment, pObject, PI / 2, PI / 2, 0, m_dVesselTantoTan, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd2TypeFND(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd2TypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd2TypeFND(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd2TypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd2TypeFND(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd2TypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd2TypeFND(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd2TypeFND"
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
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd2TypeFND(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd2TypeFND"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd2Type = EndType_FnD Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd2TypeFND(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd2TypeFND"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End2TypeFLGD
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd2TypeFLGD(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd2TypeFLGD"
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
    TransformAndRotate oDesignEquipment, pObject, PI / 2, PI / 2, 0, m_dVesselTantoTan, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd2TypeFLGD(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd2TypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd2TypeFLGD(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd2TypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject
    
    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadFlangedThick1").Value = m_dEndHeadFlangedThick1
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadFlangedThick2").Value = m_dEndHeadFlangedThick2
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
       
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd2TypeFLGD(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd2TypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd2TypeFLGD(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd2TypeFLGD"
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
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd2TypeFLGD(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd2TypeFLGD"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd2Type = EndType_FLGD Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd2TypeFLGD(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd2TypeFLGD"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End2TypeHemi
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd2TypeHemi(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd2TypeHemi"
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
    TransformAndRotate oDesignEquipment, pObject, PI / 2, PI / 2, 0, m_dVesselTantoTan, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd2TypeHemi(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd2TypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd2TypeHemi(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd2TypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
  
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd2TypeHemi(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd2TypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd2TypeHemi(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd2TypeHemi"
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
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd2TypeHemi(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd2TypeHemi"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd2Type = EndType_HEMI Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd2TypeHemi(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd2TypeHemi"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End2TypeTORC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd2TypeTORC(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd2TypeTORC"
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
    TransformAndRotate oDesignEquipment, pObject, 0, -PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd2TypeTORC(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd2TypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd2TypeTORC(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd2TypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeTopDiameter").Value = m_dEndHeadConeTopDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadConeHeight").Value = m_dEndHeadConeHeight
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadKnuckleRadius").Value = m_dEndHeadKnuckleRadius
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd2TypeTORC(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd2TypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd2TypeTORC(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd2TypeTORC"
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
    TransformAndRotate oEquipment, oEqpComponent, 0, -PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd2TypeTORC(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd2TypeTORC"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd2Type = EndType_TORC Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd2TypeTORC(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd2TypeTORC"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End2TypeTORS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd2TypeTORS(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd2TypeTORS"
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
    TransformAndRotate oDesignEquipment, pObject, 0, -PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd2TypeTORS(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd2TypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd2TypeTORS(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd2TypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadSphericalRadius").Value = m_dEndHeadSphericalRadius
    oAttribs.CollectionOfAttributes("IJUAEquipmentEnd").Item("EndHeadKnuckleRadius").Value = m_dEndHeadKnuckleRadius
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd2TypeTORS(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd2TypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd2TypeTORS(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd2TypeTORS"
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
    TransformAndRotate oEquipment, oEqpComponent, 0, -PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd2TypeTORS(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd2TypeTORS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd2Type = EndType_TORS Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd2TypeTORS(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd2TypeTORS"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Custom Methods For End2TypeFLAT
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEnd2TypeFLAT(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructEnd2TypeFLAT"
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
    TransformAndRotate oDesignEquipment, pObject, PI / 2, PI / 2, 0, m_dVesselTantoTan, 0, 0
    
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEnd2TypeFLAT(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEnd2TypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEnd2TypeFLAT(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEnd2TypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAVesselDiameter").Item("VesselDiameter").Value = m_dVesselDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEnd2TypeFLAT(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEnd2TypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryEnd2TypeFLAT(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryEnd2TypeFLAT"
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
    TransformAndRotate oEquipment, oEqpComponent, PI / 2#, PI / 2#, 0, m_dVesselTantoTan, 0, 0
    
    Set oEquipment = Nothing
    Set oEqpComponent = Nothing
    Set oEqpCompEquip = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEnd2TypeFLAT(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalEnd2TypeFLAT"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dEnd2Type = EndType_FLAT Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseEnd2TypeFLAT(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseEnd2TypeFLAT"
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

