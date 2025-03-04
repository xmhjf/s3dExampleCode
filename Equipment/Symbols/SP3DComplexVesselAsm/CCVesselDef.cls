VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCVesselDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"CAD"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:22:22 PM"
Attribute VB_Ext_KEY = "SP3DEqpCADTemplateVersion" ,"1.2"
'******************************************************************************
' Copyright (C) 2004, Intergraph Corporation. All rights reserved.
'
'File
'   CCVesselDef.cls
'
'Author
'       BBK
'
'Description
'       Defines the behavior of members of an Equipment that make a custom assembly
'       When a symbol is used with the smart occurrence that contain rigid geometry,
'       This class usually, but is not limited to, manages the nozzles of the equipment.
'Notes
'   From Template Version:1.2
'
'History:
'   02.Jan.2005    V6UpgradeSO     Initial Creation
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'   04.Jun.2005    MS  DI78930  Modified Eqp/EqpComp Symbols to use new methods of Equipment CAD helper
'   16.June.2005   MS TR79730 Some issues with the deletion of members is resolved;made datumshapes/foundationports can be deleted etc
'   24.June.2005   MS TR80576
'   08.SEP.2006    KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   04.Oct.2006    RRK  TR-105997 Corrected the NozzleIndex values in the argument list of CreateNozzleGivenIndex function call for NozzleC29 and NozzleH16
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SP3DComplexVesselAsm:CCVesselDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors

' Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt
Private m_avSymbolArrayOfInputs()   As Variant
Private PI           As Double



Private m_dSupportHeight1    As Double
Private m_dSupportDiameter1  As Double
Private m_dDomeHeight1  As Double
Private m_dVesselHeight1  As Double
Private m_dVesselDiameter1  As Double
Private m_dSupportHeight2  As Double
Private m_dDomeHeight2  As Double
Private m_dVesselHeight2  As Double
Private m_dVesselDiameter2  As Double
Private m_dSupportHeight3  As Double
Private m_dVesselHeight3  As Double
Private m_dVesselDiameter3  As Double
Private m_dSupportHeight4  As Double
Private m_dVesselHeight4  As Double
Private m_dVesselDiameter4  As Double
Private m_dDomeHeight4  As Double


Private m_oNorth                    As IJDVector
Private m_oEast                     As IJDVector
Private m_oElevation                As IJDVector

Private Sub Class_Initialize()
   PI = 4 * Atn(1)
 Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
    
    m_oEquipCADHelper.ProjectName = "SP3DComplexVesselAsm"
    m_oEquipCADHelper.ClassName = "CCVesselDef"
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

    'Add new member(NozzleG1) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleG1", 3, "CMConstructNozzleG1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleG1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleG1"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleG1"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleG1"
    'Add properties for (NozzleG1)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleG1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleG1", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleG1GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleG1", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleE12) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleE12", 4, "CMConstructNozzleE12", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleE12"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleE12"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleE12"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleE12"
    'Add properties for (NozzleE12)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleE12Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleE12", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleE12GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleE12", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleB3) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleB3", 5, "CMConstructNozzleB3", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleB3"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleB3"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleB3"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleB3"
    'Add properties for (NozzleB3)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleB3Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleB3", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleB3GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleB3", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleE24) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleE24", 6, "CMConstructNozzleE24", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleE24"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleE24"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleE24"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleE24"
    'Add properties for (NozzleE24)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleE24Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleE24", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleE24GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleE24", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleC5) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleC5", 7, "CMConstructNozzleC5", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleC5"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleC5"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleC5"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleC5"
    'Add properties for (NozzleC5)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleC5Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleC5", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleC5GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleC5", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleA6) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleA6", 8, "CMConstructNozzleA6", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleA6"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleA6"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleA6"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleA6"
    'Add properties for (NozzleA6)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleA6Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleA6", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleA6GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleA6", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleD7) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleD7", 9, "CMConstructNozzleD7", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleD7"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleD7"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleD7"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleD7"
    'Add properties for (NozzleD7)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleD7Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleD7", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleD7GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleD7", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleC18) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleC18", 10, "CMConstructNozzleC18", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleC18"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleC18"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleC18"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleC18"
    'Add properties for (NozzleC18)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleC18Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleC18", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleC18GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleC18", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleC29) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleC29", 11, "CMConstructNozzleC29", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleC29"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleC29"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleC29"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleC29"
    'Add properties for (NozzleC29)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleC29Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleC29", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleC29GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleC29", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleL110) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleL110", 12, "CMConstructNozzleL110", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleL110"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleL110"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleL110"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleL110"
    'Add properties for (NozzleL110)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleL110Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleL110", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleL110GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleL110", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleL211) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleL211", 13, "CMConstructNozzleL211", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleL211"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleL211"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleL211"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleL211"
    'Add properties for (NozzleL211)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleL211Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleL211", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleL211GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleL211", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleN12) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleN12", 14, "CMConstructNozzleN12", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleN12"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleN12"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleN12"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleN12"
    'Add properties for (NozzleN12)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleN12Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleN12", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleN12GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleN12", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleM13) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleM13", 15, "CMConstructNozzleM13", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleM13"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleM13"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleM13"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleM13"
    'Add properties for (NozzleM13)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleM13Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleM13", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleM13GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleM13", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleP14) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleP14", 16, "CMConstructNozzleP14", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleP14"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleP14"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleP14"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleP14"
    'Add properties for (NozzleP14)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleP14Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleP14", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleP14GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleP14", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleJ15) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleJ15", 17, "CMConstructNozzleJ15", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleJ15"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleJ15"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleJ15"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleJ15"
    'Add properties for (NozzleJ15)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleJ15Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleJ15", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleJ15GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleJ15", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleH16) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleH16", 18, "CMConstructNozzleH16", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleH16"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleH16"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleH16"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleH16"
    'Add properties for (NozzleH16)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleH16Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleH16", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleH16GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleH16", imsCOOKIE_ID_USS_LIB
    
        
    Set oAggregatorDescription = Nothing
    Set oMemberDescriptions = Nothing
    Set oMemberDescription = Nothing
    Set oPropertyDescriptions = Nothing
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation"
    LogCalls METHOD
    
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
        
    Select Case oMemberDescription.name
        Case "DP1", "DP2", "NozzleG1", "NozzleE12", "NozzleB3", "NozzleE24", "NozzleC5", "NozzleA6", "NozzleD7", "NozzleC18", "NozzleC29", "NozzleL110", "NozzleL211", "NozzleN12", "NozzleM13", "NozzleP14", "NozzleJ15", "NozzleH16"
            Select Case UCase(pAttrToChange.InterfaceName)
                Case "IJDELETABLEMEMBER"
                    If UCase(pAttrToChange.AttrName) = "CANBEDELETED" Then
                        m_oEquipCADHelper.MakeMemberModifyable oMemberDescription, pIJDAttrs, CBool(varNewAttrValue)
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
    
    Select Case oMemberDescription.name
        Case "DP1", "DP2"
    '   m_oEquipCADHelper.SetMemberPropertiesAsReadOnly oMemberDescription, CollAllDisplayedValues, False
    
            Dim oAttrCollection As Collection
            Dim oAttributeDescriptor As IJEquipAttrDescriptor
            Dim m As Long
            
            Set oAttrCollection = CollAllDisplayedValues
        
            For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                    Case "IJUADATUMSHAPE"
                        oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                    Case Else
                        '
                End Select
            Next
        
            Set oAttrCollection = Nothing
        Case Else
            '
    End Select
    
        Set oMemberDescription = Nothing
    
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
'    Dim dEqpX As Double, dEqpY As Double, dEqpZ As Double
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
' Custom Methods for NozzleG1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleG1(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleG1"
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
    Dim TotalHeightofAssm As Double
    TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4
    oOrientation.N1 = -TotalHeightofAssm + m_dSupportHeight1 - m_dDomeHeight1 - oNozzle.Length * 0.8
    oOrientation.N2 = 0
    oOrientation.OR1 = 0
    
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleG1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleG1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleG1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleG1"
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
        If oMemberDesc.name = "DP2" Then
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

Public Sub CMEvaluateNozzleG1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleG1"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleG1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleG1"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
     Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
     Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
   
    'checking whether the canbemodified attribute changed or not
     If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
        'creating nozzle
         'Create the nozzle orientation and set it on the nozzle
         Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
         oOrientation.PlacementType = Axial
         Dim TotalHeightofAssm As Double
         TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4
         oOrientation.N1 = -TotalHeightofAssm + m_dSupportHeight1 - m_dDomeHeight1 - oNozzle.Length * 0.8
         oOrientation.N2 = 0
         oOrientation.OR1 = 0
         Set oNozzle = Nothing
         Set oOrientation = Nothing
     End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleG1(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleG1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleG1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleG1"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleE12
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleE12(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleE12"
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
    oOrientation.N1 = m_dDomeHeight4 + oNozzle.Length * 0.8
    oOrientation.N2 = 0
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleE12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleE12"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleE12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleE12"
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
        If oMemberDesc.name = "DP2" Then
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

Public Sub CMEvaluateNozzleE12(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleE12"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleE12(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleE12"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
     Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
     Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
    
    'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
            'createnozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
            oOrientation.PlacementType = Axial
            oOrientation.N1 = m_dDomeHeight4 + oNozzle.Length * 0.8
            oOrientation.N2 = 0
            oOrientation.OR1 = 0
            Set oNozzle = Nothing
            Set oOrientation = Nothing 'Transform the nozzle so that it behaves like a rigid body inside the equipment
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleE12(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleE12"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleE12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleE12"
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
    oOrientation.PlacementType = Radial
    oOrientation.N1 = -m_dVesselHeight2 - m_dSupportHeight3 - m_dVesselHeight3 - m_dSupportHeight4 - m_dVesselHeight4 - m_dDomeHeight4 - oNozzle.Length * 0.8
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
        If oMemberDesc.name = "DP2" Then
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
    
    'getting the IJDMemberDescription from IJDPropertyDescription
    Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
    Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
   
     'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
            ' create nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            oOrientation.N1 = -m_dVesselHeight2 - m_dSupportHeight3 - m_dVesselHeight3 - m_dSupportHeight4 - m_dVesselHeight4 - m_dDomeHeight4 - oNozzle.Length * 0.8
            oOrientation.N2 = 0
            oOrientation.OR1 = 0
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
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
' Custom Methods for NozzleE24
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleE24(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleE24"
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
    Dim TotalHeightofAssm As Double
    TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4 + m_dDomeHeight4
    
    oOrientation.N1 = (TotalHeightofAssm) - m_dDomeHeight4 - 0.5 * m_dVesselHeight4
    oOrientation.N2 = m_dVesselDiameter4 / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleE24(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleE24"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleE24(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleE24"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleE24(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleE24"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleE24(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleE24"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
    Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
    Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
          
     'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
             'create nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            Dim TotalHeightofAssm As Double
            TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4 + m_dDomeHeight4
            oOrientation.N1 = (TotalHeightofAssm) - m_dDomeHeight4 - 0.5 * m_dVesselHeight4
            oOrientation.N2 = m_dVesselDiameter4 / 2 + oNozzle.Length * 0.8
            oOrientation.OR1 = 0
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleE24(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleE24"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleE24(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleE24"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleC5
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleC5(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleC5"
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
    Dim TotalHeightofAssm As Double
    TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4 + m_dDomeHeight4
    oOrientation.N1 = (TotalHeightofAssm) - m_dDomeHeight4 - m_dVesselHeight4 - m_dSupportHeight4 - 0.5 * m_dVesselHeight3
    oOrientation.N2 = m_dVesselDiameter3 / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleC5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleC5"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleC5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleC5"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleC5(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleC5"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleC5(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleC5"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
    Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
     Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
    
    'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
            'create nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
            oOrientation.PlacementType = Radial
            Dim TotalHeightofAssm As Double
            TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4 + m_dDomeHeight4
            oOrientation.N1 = (TotalHeightofAssm) - m_dDomeHeight4 - m_dVesselHeight4 - m_dSupportHeight4 - 0.5 * m_dVesselHeight3
            oOrientation.N2 = m_dVesselDiameter3 / 2 + oNozzle.Length * 0.8
            oOrientation.OR1 = PI / 2
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
     Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleC5(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleC5"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleC5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleC5"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleA6
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleA6(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleA6"
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
    
    oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + 0.5 * m_dVesselHeight2
    oOrientation.N2 = (m_dVesselDiameter2 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = PI

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleA6(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleA6"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleA6(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleA6"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleA6(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleA6"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleA6(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleA6"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
    Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
    Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
    
    'checking whether the canbemodified attribute changed or not
      If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
          'Create Nozzle
          'Create the nozzle orientation and set it on the nozzle
          Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
          oOrientation.PlacementType = Radial
          oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + 0.5 * m_dVesselHeight2
          oOrientation.N2 = (m_dVesselDiameter2 / 2 + oNozzle.Length * 0.8)
          oOrientation.OR1 = PI
          Set oNozzle = Nothing
          Set oOrientation = Nothing
      
      End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleA6(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleA6"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleA6(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleA6"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleD7
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleD7(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleD7"
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
    oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + 0.5 * m_dVesselHeight2
    oOrientation.N2 = (m_dVesselDiameter2 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleD7(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleD7"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleD7(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleD7"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleD7(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleD7"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleD7(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleD7"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
     Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
    Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
    
    'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
            'creating nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + 0.5 * m_dVesselHeight2
            oOrientation.N2 = (m_dVesselDiameter2 / 2 + oNozzle.Length * 0.8)
            oOrientation.OR1 = 0
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
     Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleD7(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleD7"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleD7(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleD7"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleC18
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleC18(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleC18"
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
    oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + 0.5 * m_dVesselHeight2
    oOrientation.N2 = (m_dVesselDiameter2 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = 3 * PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleC18(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleC18"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleC18(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleC18"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleC18(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleC18"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleC18(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleC18"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
    Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
     Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
    
    'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
            'create nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + 0.5 * m_dVesselHeight2
            oOrientation.N2 = (m_dVesselDiameter2 / 2 + oNozzle.Length * 0.8)
            oOrientation.OR1 = 3 * PI / 2
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleC18(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleC18"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleC18(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleC18"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleC29
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleC29(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleC29"
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
    oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + 0.5 * m_dVesselHeight2
    oOrientation.N2 = (m_dVesselDiameter2 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleC29(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleC29"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleC29(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleC29"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleC29(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleC29"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleC29(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleC29"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
    Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
    Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
     Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
      
    'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
            'create nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + 0.5 * m_dVesselHeight2
            oOrientation.N2 = (m_dVesselDiameter2 / 2 + oNozzle.Length * 0.8)
            oOrientation.OR1 = PI / 2
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleC29(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleC29"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleC29(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleC29"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleL110
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleL110(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleL110"
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
    oOrientation.PlacementType = Radial
    oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 * 0.25
    oOrientation.N2 = (m_dVesselDiameter1 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleL110(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleL110"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleL110(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleL110"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleL110(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleL110"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleL110(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleL110"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
     Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
     Set oNozzle = oPropertyDescription.object
     Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
          
    'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
            'creating nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 * 0.25
            oOrientation.N2 = (m_dVesselDiameter1 / 2 + oNozzle.Length * 0.8)
            oOrientation.OR1 = PI / 2
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleL110(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleL110"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleL110(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleL110"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleL211
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleL211(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleL211"
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
    oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 * 0.75
    oOrientation.N2 = (m_dVesselDiameter1 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleL211(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleL211"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleL211(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleL211"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleL211(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleL211"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleL211(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleL211"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
     Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
     Set oNozzle = oPropertyDescription.object
     Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
          
     'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
            'creating nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 * 0.75
            oOrientation.N2 = (m_dVesselDiameter1 / 2 + oNozzle.Length * 0.8)
            oOrientation.OR1 = PI / 2
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleL211(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleL211"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleL211(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleL211"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleN12
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleN12(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleN12"
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
    oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 * 0.25
    oOrientation.N2 = (m_dVesselDiameter1 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = 3 * PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleN12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleN12"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleN12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleN12"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleN12(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleN12"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleN12(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleN12"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
     Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
     Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
          
    'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
             'creating nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 * 0.25
            oOrientation.N2 = (m_dVesselDiameter1 / 2 + oNozzle.Length * 0.8)
            oOrientation.OR1 = 3 * PI / 2
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleN12(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleN12"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleN12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleN12"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleM13
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleM13(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleM13"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 13, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 * 0.75
    oOrientation.N2 = (m_dVesselDiameter1 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = 3 * PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleM13(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleM13"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleM13(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleM13"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleM13(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleM13"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleM13(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleM13"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
     Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
    Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
          
    'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
            'creating nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 * 0.75
            oOrientation.N2 = (m_dVesselDiameter1 / 2 + oNozzle.Length * 0.8)
            oOrientation.OR1 = 3 * PI / 2
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleM13(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleM13"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleM13(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleM13"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleP14
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleP14(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleP14"
    On Error GoTo ErrorHandler
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 14, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 * 0.2
    oOrientation.N2 = (m_dVesselDiameter1 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleP14(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleP14"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleP14(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleP14"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleP14(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleP14"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleP14(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleP14"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
     Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
     Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
          
     'checking the canbemodified attribute
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
             'creating nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            oOrientation.N1 = m_dSupportHeight1 + m_dVesselHeight1 * 0.2
            oOrientation.N2 = (m_dVesselDiameter1 / 2 + oNozzle.Length * 0.8)
            oOrientation.OR1 = 0
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleP14(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleP14"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleP14(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleP14"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleJ15
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleJ15(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleJ15"
    On Error GoTo ErrorHandler
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 15, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    Dim TotalHeightofAssm As Double
    TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4 + m_dDomeHeight4
    oOrientation.N1 = (TotalHeightofAssm) - m_dDomeHeight4 - m_dVesselHeight4 - m_dSupportHeight4 - 0.5 * m_dVesselHeight3
    oOrientation.N2 = (m_dVesselDiameter3 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = 3 * PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleJ15(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleJ15"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleJ15(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleJ15"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleJ15(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleJ15"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleJ15(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleJ15"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
     Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
     Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
          
     'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
            'creating nozzle
            'Create the nozzle orientation and set it on the nozzle
            Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            Dim TotalHeightofAssm As Double
            TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4 + m_dDomeHeight4
            oOrientation.N1 = (TotalHeightofAssm) - m_dDomeHeight4 - m_dVesselHeight4 - m_dSupportHeight4 - 0.5 * m_dVesselHeight3
            oOrientation.N2 = (m_dVesselDiameter3 / 2 + oNozzle.Length * 0.8)
            oOrientation.OR1 = 3 * PI / 2
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleJ15(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleJ15"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleJ15(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleJ15"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleH16
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleH16(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleH16"
    On Error GoTo ErrorHandler
        
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 16, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    Dim TotalHeightofAssm As Double
    TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4 + m_dDomeHeight4
    oOrientation.N1 = (TotalHeightofAssm) - m_dDomeHeight4 - m_dVesselHeight4 - m_dSupportHeight4 - 0.5 * m_dVesselHeight3
    oOrientation.N2 = (m_dVesselDiameter3 / 2 + oNozzle.Length * 0.8)
    oOrientation.OR1 = 0
        

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleH16(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleH16"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleH16(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleH16"
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
        If oMemberDesc.name = "DP1" Then
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

Public Sub CMEvaluateNozzleH16(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleH16"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleH16(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleH16"
    On Error GoTo ErrorHandler
    
    'getting the IJDMemberDescription from IJDPropertyDescription
     Dim omembdes As IJDMemberDescription
    Dim oattrib As IJDAttributes
    Dim oNozzle As IJDNozzle
    Dim oOrientation As IJNozzleOrientation
     Set oNozzle = oPropertyDescription.object
    Set oattrib = oNozzle
    Set omembdes = m_oEquipCADHelper.GetMemberDescriptionFromChild(oattrib)
     
     'checking whether the canbemodified attribute changed or not
        If m_oEquipCADHelper.CheckMemberModifyable(omembdes) = False Then
             'creating nozzle
            'Create the nozzle orientation and set it on the nozzle
             Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, pObject)
            oOrientation.PlacementType = Radial
            Dim TotalHeightofAssm As Double
            TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4 + m_dDomeHeight4
            oOrientation.N1 = (TotalHeightofAssm) - m_dDomeHeight4 - m_dVesselHeight4 - m_dSupportHeight4 - 0.5 * m_dVesselHeight3
            oOrientation.N2 = (m_dVesselDiameter3 / 2 + oNozzle.Length * 0.8)
            oOrientation.OR1 = 0
            Set oNozzle = Nothing
            Set oOrientation = Nothing
        End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleH16(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleH16"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleH16(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleH16"
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
    m_dSupportHeight1 = m_avSymbolArrayOfInputs(2)
    m_dSupportDiameter1 = m_avSymbolArrayOfInputs(3)
    m_dDomeHeight1 = m_avSymbolArrayOfInputs(4)
    m_dVesselHeight1 = m_avSymbolArrayOfInputs(5)
    m_dVesselDiameter1 = m_avSymbolArrayOfInputs(6)
    m_dSupportHeight2 = m_avSymbolArrayOfInputs(7)
    m_dDomeHeight2 = m_avSymbolArrayOfInputs(8)
    m_dVesselHeight2 = m_avSymbolArrayOfInputs(9)
    m_dVesselDiameter2 = m_avSymbolArrayOfInputs(10)
    m_dSupportHeight3 = m_avSymbolArrayOfInputs(11)
    m_dVesselHeight3 = m_avSymbolArrayOfInputs(12)
    m_dVesselDiameter3 = m_avSymbolArrayOfInputs(13)
    m_dSupportHeight4 = m_avSymbolArrayOfInputs(14)
    m_dVesselHeight4 = m_avSymbolArrayOfInputs(15)
    m_dVesselDiameter4 = m_avSymbolArrayOfInputs(16)
    m_dDomeHeight4 = m_avSymbolArrayOfInputs(17)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub HandleError(sModule As String, sMethod As String)

    If Not m_oEditErrors Is Nothing Then
        m_oEditErrors.AddFromErr Err, "", sMethod, sModule
    End If

    Err.Raise Err.Number, Err.Source & " " & sMethod, Err.description, _
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

Private Sub PositionAndOrientDP2(Equipment As IJEquipment, Shape As IJShape)

    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
     Dim TotalHeightofAssm As Double
    TotalHeightofAssm = m_dSupportHeight1 + m_dVesselHeight1 + m_dSupportHeight2 + m_dVesselHeight2 + m_dSupportHeight3 + m_dVesselHeight3 + m_dSupportHeight4 + m_dVesselHeight4
 
    oPosition.Set 0, 0, TotalHeightofAssm
    
    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oElevation, m_oNorth
    
    Set oPosition = Nothing

End Sub
