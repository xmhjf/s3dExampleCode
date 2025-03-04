VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VDrumWithLegsDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"CAD"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:22:39 PM"
Attribute VB_Ext_KEY = "SP3DEqpCADTemplateVersion" ,"1.2"
'******************************************************************************
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'File
'   VDrumWithLegsDef.cls
'
'Author
'       KB
'
'Description
'       Defines the behavior of members of an Equipment that make a custom assembly
'       When a symbol is used with the smart occurrence that contain rigid geometry,
'       This class usually, but is not limited to, manages the nozzles of the equipment.
'Notes
'   From Template Version:1.2
'
'History:
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'   4.Oct.2006      RRK                     TR105997 Corrected the NozzleIndex from 1 to 4 in the argument list of CreateNozzleGivenIndex function for NozzleH14
'******************************************************************************
Option Explicit

Private Const MODULE = "SP3DVerDrumWiLegsAsm:VDrumWithLegsDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors

' Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt
Private m_avSymbolArrayOfInputs()   As Variant

Private PI As Double

Private m_dVesselDiameter As Double
Private m_dVesselTantoTan As Double
Private m_dCPtoSupport As Double
Private m_dInsulationThickness As Double
Private m_dSupportOrientation As Double
Private m_dSupportLength As Double
Private m_dSupportWidth As Double
Private m_dSupportHeight As Double
Private m_dBackingPlateThickness As Double
Private m_dBackingPlateDepth As Double
Private m_dLengthBaseplate As Double
Private m_dWidthBaseplate As Double
Private m_dBasePlateThickness As Double


Private m_oNorth                    As IJDVector
Private m_oEast                     As IJDVector
Private m_oElevation                As IJDVector

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
    
    m_oEquipCADHelper.ProjectName = "SP3DVerDrumWiLegsAsm"
    m_oEquipCADHelper.ClassName = "VDrumWithLegsDef"
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
End Sub

Private Sub Class_Terminate()
    LogCalls "Class_Terminate"
    Set m_oNorth = Nothing
    Set m_oEast = Nothing
    Set m_oElevation = Nothing

'    Set m_oPartFclt = Nothing

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

    'Add new member(NozzleA11) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleA11", 2, "CMConstructNozzleA11", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleA11"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleA11"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleA11"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleA11"
    'Add properties for (NozzleA11)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleA11Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleA11", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleA11GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleA11", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleA22) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleA22", 3, "CMConstructNozzleA22", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleA22"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleA22"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleA22"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleA22"
    'Add properties for (NozzleA22)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleA22Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleA22", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleA22GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleA22", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleB3) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleB3", 4, "CMConstructNozzleB3", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleB3"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleB3"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleB3"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleB3"
    'Add properties for (NozzleB3)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleB3Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleB3", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleB3GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleB3", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleH14) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleH14", 5, "CMConstructNozzleH14", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleH14"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleH14"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleH14"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleH14"
    'Add properties for (NozzleH14)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleH14Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleH14", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleH14GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleH14", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleH25) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleH25", 6, "CMConstructNozzleH25", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleH25"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleH25"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleH25"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleH25"
    'Add properties for (NozzleH25)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleH25Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleH25", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleH25GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleH25", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleL16) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleL16", 7, "CMConstructNozzleL16", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleL16"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleL16"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleL16"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleL16"
    'Add properties for (NozzleL16)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleL16Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleL16", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleL16GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleL16", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleL27) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleL27", 8, "CMConstructNozzleL27", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleL27"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleL27"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleL27"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleL27"
    'Add properties for (NozzleL27)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleL27Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleL27", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleL27GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleL27", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleL38) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleL38", 9, "CMConstructNozzleL38", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleL38"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleL38"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleL38"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleL38"
    'Add properties for (NozzleL38)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleL38Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleL38", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleL38GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleL38", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleL49) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleL49", 10, "CMConstructNozzleL49", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleL49"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleL49"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleL49"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleL49"
    'Add properties for (NozzleL49)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleL49Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleL49", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleL49GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleL49", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleS10) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleS10", 11, "CMConstructNozzleS10", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleS10"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleS10"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleS10"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleS10"
    'Add properties for (NozzleS10)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleS10Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleS10", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleS10GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleS10", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleT11) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleT11", 12, "CMConstructNozzleT11", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleT11"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleT11"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleT11"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleT11"
    'Add properties for (NozzleT11)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleT11Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleT11", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleT11GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleT11", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleV12) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleV12", 13, "CMConstructNozzleV12", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleV12"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleV12"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleV12"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleV12"
    'Add properties for (NozzleV12)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleV12Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleV12", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleV12GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleV12", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleEqpFoundationPort13) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleEqpFoundationPort13", 14, "CMConstructNozzleEqpFoundationPort13", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleEqpFoundationPort13"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleEqpFoundationPort13"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleEqpFoundationPort13"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleEqpFoundationPort13"
    'Add properties for (NozzleEqpFoundationPort13)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort13Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleEqpFoundationPort13", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort13GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleEqpFoundationPort13", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleEqpFoundationPort214) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleEqpFoundationPort214", 15, "CMConstructNozzleEqpFoundationPort214", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleEqpFoundationPort214"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleEqpFoundationPort214"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleEqpFoundationPort214"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleEqpFoundationPort214"
    'Add properties for (NozzleEqpFoundationPort214)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort214Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleEqpFoundationPort214", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort214GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleEqpFoundationPort214", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleEqpFoundationPort315) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleEqpFoundationPort315", 16, "CMConstructNozzleEqpFoundationPort315", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleEqpFoundationPort315"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleEqpFoundationPort315"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleEqpFoundationPort315"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleEqpFoundationPort315"
    'Add properties for (NozzleEqpFoundationPort315)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort315Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleEqpFoundationPort315", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort315GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleEqpFoundationPort315", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleEqpFoundationPort416) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleEqpFoundationPort416", 17, "CMConstructNozzleEqpFoundationPort416", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleEqpFoundationPort416"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleEqpFoundationPort416"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleEqpFoundationPort416"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleEqpFoundationPort416"
    'Add properties for (NozzleEqpFoundationPort416)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort416Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleEqpFoundationPort416", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort416GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleEqpFoundationPort416", imsCOOKIE_ID_USS_LIB
    
        
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
    Const METHOD = "IJDUserSymbolServices_EditOccurence"
    LogCalls METHOD
    'Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
    IJDUserSymbolServices_EditOccurence = False

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation of IJEquipUserAttrMgmt
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IJEquipUserAttrMgmt_OnAttributeChange(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As IJEquipAttrDescriptor, ByVal varNewAttrValue As Variant) As String
    Const METHOD = "IJEquipUserAttrMgmt_OnAttributeChange"
    On Error GoTo ErrorHandler
    
    
    Dim oMemberDescription As IJDMemberDescription
    
    
    Set oMemberDescription = m_oEquipCADHelper.GetMemberDescriptionFromChild(pIJDAttrs)
        
    Select Case oMemberDescription.Name
        Case "DP1", "NozzleA11", "NozzleA22", "NozzleB3", "NozzleH14", "NozzleH25", "NozzleL16", "NozzleL27", "NozzleL38", "NozzleL49", "NozzleS10", "NozzleT11", "NozzleV12", "NozzleEqpFoundationPort13", "NozzleEqpFoundationPort214", "NozzleEqpFoundationPort315", "NozzleEqpFoundationPort416"
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
        Case "DP1"
'            m_oEquipCADHelper.SetMemberPropertiesAsReadOnly oMemberDescription, CollAllDisplayedValues, False
            For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                    Case "IJUADATUMSHAPE"
                        oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                    Case Else
                        '
                End Select
            Next
        

        Case Else
            '
    End Select
    
    Select Case oMemberDescription.Name
        Case "NozzleB3", "NozzleH14", "NozzleH25", "NozzleL16"
      
            For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                Case "IJDELETABLEMEMBER"
                    If UCase(oAttributeDescriptor.AttrName) = "CANBEDELETED" Then
                        oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                    End If
                Case "IJNOZZLEORIENTATION"
                    'The nozzle is placed by Point
                    oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                Case Else
                        '
                End Select
            Next
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
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructEquipment"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMSetInputsEquipment"
    On Error GoTo ErrorHandler
    
    m_oEquipCADHelper.SetSmartItemAsInputToSymbol pAggregatorDescription
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMRemoveInputsEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMRemoveInputsEquipment"
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
' Custom Methods for NozzleA11
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleA11(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleA11"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 1, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    oOrientation.N1 = m_dVesselTantoTan + m_dVesselDiameter / 8 + oNozzle.Length
    oOrientation.N2 = m_dVesselDiameter / 4
    oOrientation.OR1 = PI

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleA11(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleA11"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleA11(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleA11"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleA11(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleA11"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleA11(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleA11"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleA11(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleA11"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleA11(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleA11"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleA22
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleA22(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleA22"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 2, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    
    oOrientation.N1 = m_dVesselTantoTan * 0.8
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = 3 * PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleA22(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleA22"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleA22(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleA22"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleA22(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleA22"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleA22(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleA22"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleA22(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleA22"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleA22(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleA22"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleB3
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleB3(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleB3"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 3, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    
    oOrientation.N1 = -m_dVesselDiameter / 4 - oNozzle.Length * 0.8
    oOrientation.N2 = 0
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleB3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleB3"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleB3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleB3"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleB3(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleB3"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleB3(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleB3"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleB3(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleB3"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleB3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleB3"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleH14
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleH14(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleH14"
    On Error GoTo ErrorHandler
        
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
     m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 4, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    
    oOrientation.N1 = m_dVesselTantoTan * 0.3
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = 7 * PI / 4

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleH14(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleH14"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleH14(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleH14"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleH14(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleH14"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleH14(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleH14"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleH14(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleH14"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleH14(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleH14"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleH25
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleH25(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleH25"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 5, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    
    oOrientation.N1 = m_dVesselTantoTan * 0.8
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleH25(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleH25"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleH25(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleH25"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleH25(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleH25"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleH25(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleH25"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleH25(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleH25"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleH25(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleH25"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleL16
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleL16(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleL16"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 6, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    
    oOrientation.N1 = m_dVesselTantoTan * 0.8
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = 5 * PI / 4

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleL16(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleL16"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleL16(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleL16"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleL16(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleL16"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleL16(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleL16"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleL16(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleL16"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleL16(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleL16"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleL27
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleL27(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleL27"
    On Error GoTo ErrorHandler
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 7, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    
    oOrientation.N1 = m_dVesselTantoTan * 0.3
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = 5 * PI / 4

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleL27(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleL27"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleL27(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleL27"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleL27(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleL27"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleL27(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleL27"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleL27(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleL27"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleL27(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleL27"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleL38
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleL38(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleL38"
    On Error GoTo ErrorHandler
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 8, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    
    oOrientation.N1 = m_dVesselTantoTan * 0.8
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = 3 * PI / 4

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleL38(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleL38"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleL38(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleL38"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleL38(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleL38"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleL38(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleL38"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleL38(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleL38"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleL38(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleL38"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleL49
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleL49(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleL49"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 9, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    
    oOrientation.N1 = m_dVesselTantoTan * 0.3
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = 3 * PI / 4

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleL49(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleL49"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleL49(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleL49"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleL49(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleL49"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleL49(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleL49"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleL49(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleL49"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleL49(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleL49"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleS10
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleS10(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleS10"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 10, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    
    oOrientation.N1 = m_dVesselTantoTan + m_dVesselDiameter / 8 + oNozzle.Length
    oOrientation.N2 = m_dVesselDiameter / 4
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleS10(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleS10"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleS10(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleS10"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleS10(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleS10"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleS10(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleS10"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleS10(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleS10"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleS10(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleS10"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleT11
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleT11(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleT11"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 11, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    oOrientation.N1 = m_dVesselTantoTan * 0.3
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = PI / 4

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleT11(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleT11"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleT11(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleT11"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleT11(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleT11"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleT11(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleT11"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleT11(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleT11"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleT11(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleT11"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleV12
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleV12(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleV12"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 12, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    
    oOrientation.N1 = m_dVesselTantoTan + m_dVesselDiameter / 4 + oNozzle.Length
    oOrientation.N2 = 0
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleV12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleV12"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleV12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleV12"
    On Error GoTo ErrorHandler
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleV12(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleV12"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleV12(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleV12"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment

        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleV12(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleV12"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleV12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleV12"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleEqpFoundationPort13
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleEqpFoundationPort13(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleEqpFoundationPort13"
    On Error GoTo ErrorHandler
        
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 13
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleEqpFoundationPort13(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleEqpFoundationPort13"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleEqpFoundationPort13(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleEqpFoundationPort13"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleEqpFoundationPort13(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleEqpFoundationPort13"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleEqpFoundationPort13(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleEqpFoundationPort13"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 13
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleEqpFoundationPort13(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleEqpFoundationPort13"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleEqpFoundationPort13(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleEqpFoundationPort13"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleEqpFoundationPort214
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleEqpFoundationPort214(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleEqpFoundationPort214"
    On Error GoTo ErrorHandler
        
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 14
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleEqpFoundationPort214(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleEqpFoundationPort214"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleEqpFoundationPort214(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleEqpFoundationPort214"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleEqpFoundationPort214(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleEqpFoundationPort214"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleEqpFoundationPort214(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleEqpFoundationPort214"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 14
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleEqpFoundationPort214(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleEqpFoundationPort214"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleEqpFoundationPort214(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleEqpFoundationPort214"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleEqpFoundationPort315
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleEqpFoundationPort315(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleEqpFoundationPort315"
    On Error GoTo ErrorHandler
        
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 15
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleEqpFoundationPort315(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleEqpFoundationPort315"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleEqpFoundationPort315(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleEqpFoundationPort315"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleEqpFoundationPort315(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleEqpFoundationPort315"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleEqpFoundationPort315(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleEqpFoundationPort315"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 15
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleEqpFoundationPort315(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleEqpFoundationPort315"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleEqpFoundationPort315(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleEqpFoundationPort315"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleEqpFoundationPort416
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleEqpFoundationPort416(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleEqpFoundationPort416"
    On Error GoTo ErrorHandler
        
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 16
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleEqpFoundationPort416(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleEqpFoundationPort416"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleEqpFoundationPort416(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleEqpFoundationPort416"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleEqpFoundationPort416(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleEqpFoundationPort416"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleEqpFoundationPort416(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleEqpFoundationPort416"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 16
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleEqpFoundationPort416(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleEqpFoundationPort416"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleEqpFoundationPort416(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleEqpFoundationPort416"
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
    'Set m_oPartFclt = m_avSymbolArrayOfInputs(1)
    
    
    m_dVesselDiameter = m_avSymbolArrayOfInputs(2)
    m_dVesselTantoTan = m_avSymbolArrayOfInputs(3)
    m_dCPtoSupport = m_avSymbolArrayOfInputs(4)
    m_dInsulationThickness = m_avSymbolArrayOfInputs(5)
    m_dSupportOrientation = m_avSymbolArrayOfInputs(6)
    m_dSupportLength = m_avSymbolArrayOfInputs(7)
    m_dSupportWidth = m_avSymbolArrayOfInputs(8)
    m_dSupportHeight = m_avSymbolArrayOfInputs(9)
    m_dBackingPlateThickness = m_avSymbolArrayOfInputs(10)
    m_dBackingPlateDepth = m_avSymbolArrayOfInputs(11)
    m_dLengthBaseplate = m_avSymbolArrayOfInputs(12)
    m_dWidthBaseplate = m_avSymbolArrayOfInputs(13)
    m_dBasePlateThickness = m_avSymbolArrayOfInputs(14)

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
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oElevation, m_oNorth
    
    Set oPosition = Nothing

End Sub
