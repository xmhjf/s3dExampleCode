VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHEx2ShellDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:26:30 PM"

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CHEx2ShellSym.cls
'   ProgID:         SP3DHeatEx2EndShellComp.CHEx2ShellDef
'   Author:         BBK
'   Creation Date:  Monday, Nov 01, 2003

'   This is an Exchanger Shell component Body Symbol.
'   This symbol is Fully parametric symbol, consists of 77 inputs and 10 outputs.It has 2 aspects one is Simple Physical
'   another is ReferenceGeometry. Among them 8 are physical and 2 is ReferenceGeometry. It has Five nozzles and all
'   are fully parametric.
'   This is an equipment component symbol using the Smart Occurrence instead of Part occurrence. The symbol would be of
'   two parts, one smart occurrence symbol and the other is the custom assembly definition. This file defines the symbol
'   part of the complete component and the nozzles are created in the definition(CHEx2ShellDef.cls)
'   This is a Dow Equipment E-161 Exchanger Shell symbol.
'
'   Change History:
'   dd.mmm.yyyy         who             change description
'   -----------         ---             ------------------
'
'   01.Nov.2004     BBK     Initial Creation from CHEx2EndShell.cls
'   02.Jan.2005     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const MODULE = "CHEx2ShellDef:"
Private Const m_ItemProgId As String = "SP3DHeatEx2EndShellComp.CHEx2ShellDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const GUID_CPEQUIPMENT = "{0D1E23F9-280A-4382-9370-9B0C5C363C88}"
Private Const GUID_CPEQUIPMENTCOMPONENT = "{F4FC2BB5-A507-4c59-B6E8-BC0AE3842A26}"

Private Const IJInitialPipePort = "IJInitialPipePort"
Private Const IJDynamicPipePort = "IJDynamicPipePort"
Private Const IJUAVesselDiameter = "IJUAVesselDiameter"
Private Const IJUARadialPosition = "IJUARadialPosition"
Private Const IJUAOffset = "IJUAOffset "
Private Const IJUANozzleOrientation = "IJUANozzleOrientation"
Private Const IJUANozzlePosition = "IJUANozzlePosition"

Private Const Id = "Id"
Private Const Npd = "Npd"
Private Const NPD_UNIT_TYPE = "NPDUnitType"
Private Const END_PREPARATION = "EndPreparation"
Private Const SCHEDULE_THICKNESS = "ScheduleThickness"
Private Const END_STANDARD = "EndStandard"
Private Const PRESSURE_RATING = "PressureRating"
Private Const FLOW_DIRECTION = "FlowDirection"
Private Const TERMINATION_SUB_CLASS = "TerminationSubClass"

Private Const VESSEL_DIAMETER = "VesselDiameter"
Private Const RADIAL_POSITION = "RadialPosition"
Private Const OFFSET = "Offset"
Private Const NOZZLE_ORIENTATION = "NozzleOrientation"
Private Const NOZZLE_POSITION = "NozzlePosition"
    
' Implement User Symbol Services(USS)
Implements IJDUserSymbolServices

Private m_oEqpCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors

Option Explicit

Private Sub Class_Initialize()
    Set m_oEqpCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oEqpCADHelper = Nothing
    Set m_oEditErrors = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation of IJDUserSymbolServices
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    'Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Progid should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemProgId
End Function

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                             ByVal defParams As Variant, _
                                                             ByVal pResourceMgr As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler
     
    ' This method is in charge of the creation of the symbol definition object
    Dim oSymbolDefinition As IJDSymbolDefinition
    Dim oCAFactory As GSCADCustomAssembly.IJCAFactory
    Set oCAFactory = New GSCADCustomAssembly.CAFactory
    Set oSymbolDefinition = oCAFactory.CreateCAD(pResourceMgr)
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = m_ItemProgId
    oSymbolDefinition.CodeBase = CodeBase
    
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    oSymbolDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
    
    ' Persistence behavior
    oSymbolDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    oSymbolDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
    'returned symbol definition
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    
    Dim oPropertyDescriptions As IJDPropertyDescriptions
    Dim oMemberDescriptions As IJDMemberDescriptions
    Dim oMemberDescription As IJDMemberDescription
    Dim oAD As IJDAggregatorDescription
    
    'Set up the aggregator
    Set oAD = oSymbolDefinition
    oAD.AggregatorClsid = GUID_CPEQUIPMENTCOMPONENT 'CPEquipmentComponent
    oAD.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructHeatEx2EndShellComp"
    oAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructHeatEx2EndShellComp"
    oAD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsHeatEx2EndShellComp"
    oAD.SetCMRemoveInputs imsCOOKIE_ID_USS_LIB, "CMRemoveInputsHeatEx2EndShellComp"
    
    ' Add property to aggregator
    Set oPropertyDescriptions = oAD
    oPropertyDescriptions.RemoveAll
    oPropertyDescriptions.AddProperty "HeatEx2EndShellCompProps", 1, IID_IJDATTRIBUTES, "CMEvaluateHeatEx2EndShellComp", imsCOOKIE_ID_USS_LIB
    
    ' Remove all the previous member descriptions
    Set oMemberDescriptions = oSymbolDefinition
    oMemberDescriptions.RemoveAll
    
    ' Add new member(C) to the definition
    Set oMemberDescription = oMemberDescriptions.AddMember("C", 1, "CMConstructMemberC", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsMemberC"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructMemberC"
    ' Add properties for C
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "MemberProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateMemberC", imsCOOKIE_ID_USS_LIB

    ' Add new member(D) to the definition
    Set oMemberDescription = oMemberDescriptions.AddMember("D", 2, "CMConstructMemberD", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsMemberD"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructMemberD"
    ' Add properties for D
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "MemberPropertiesD", 1, IID_IJDATTRIBUTES, "CMEvaluateMemberD", imsCOOKIE_ID_USS_LIB
    
    ' Add new member(F1) to the definition
    Set oMemberDescription = oMemberDescriptions.AddMember("F1", 3, "CMConstructMemberF1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsMemberF1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructMemberF1"
    ' Add properties for F1
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "MemberF1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateMemberF1", imsCOOKIE_ID_USS_LIB
    
    ' Add new member(F2) to the definition
    Set oMemberDescription = oMemberDescriptions.AddMember("F2", 4, "CMConstructMemberF2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsMemberF2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructMemberF2"
    ' Add properties for F2
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "MemberF2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateMemberF2", imsCOOKIE_ID_USS_LIB
    
    ' Add new member(F3) to the definition
    Set oMemberDescription = oMemberDescriptions.AddMember("F3", 5, "CMConstructMemberF3", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsMemberF3"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructMemberF3"
    ' Add properties for F3
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "MemberF3Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateMemberF3", imsCOOKIE_ID_USS_LIB
    
    Set oAD = Nothing
    Set oMemberDescriptions = Nothing
    Set oMemberDescription = Nothing
    Set oPropertyDescriptions = Nothing
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods of Aggregator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CMConstructHeatEx2EndShellComp(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMConstructHeatEx2EndShellComp"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructHeatEx2EndShellComp(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructHeatEx2EndShellComp"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsHeatEx2EndShellComp(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMSetInputsHeatEx2EndShellComp"
    On Error GoTo ErrorHandler
    
    Dim oSmartdesignobject As IJSmartDesignObject

    Set oSmartdesignobject = pAggregatorDescription.CAO

    ' This method sets SmartItem as Reference input to the symbol
    Call oSmartdesignobject.SetSmartItemAsInputToSymbol

    Set oSmartdesignobject = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMRemoveInputsHeatEx2EndShellComp(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMRemoveInputsHeatEx2EndShellComp"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateHeatEx2EndShellComp(oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateHeatEx2EndShellComp"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for MemberC
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CMConstructMemberC(ByVal pMemberDescription As IJDMemberDescription, _
                              ByVal pResourceManager As IUnknown, _
                              ByRef pObj As Object)
    Const METHOD = "CMConstructMemberC"
    On Error GoTo ErrorHandler
    
    m_oEqpCADHelper.CreateDynamicPipeNozzleFromPH pMemberDescription, pResourceManager, pObj, 1

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructMemberC(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructMemberC"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsMemberC(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsMemberC"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateMemberC(oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberC"
    On Error GoTo ErrorHandler
    
    ' Transform the nozzle so that it behaves like a rigid body inside the equipment
    'm_oEqpCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 1
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for MemberD
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CMConstructMemberD(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObj As Object)
    Const METHOD = "CMConstructMemberD"
    On Error GoTo ErrorHandler

    m_oEqpCADHelper.CreateDynamicPipeNozzleFromPH pMemberDescription, pResourceManager, pObj, 2
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructMemberD(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructMemberD"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsMemberD(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsMemberD"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateMemberD(oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberD"
    On Error GoTo ErrorHandler
    
    ' Transform the nozzle so that it behaves like a rigid body inside the equipment
    'm_oEqpCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 2
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for MemberF1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CMConstructMemberF1(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObj As Object)
    Const METHOD = "CMConstructMemberF1"
    On Error GoTo ErrorHandler

    ' Create Pipe Nozzle
    m_oEqpCADHelper.CreateDynamicPipeNozzleFromPH pMemberDescription, pResourceManager, pObj, 3

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructMemberF1(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructMemberF1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsMemberF1(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsMemberF1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateMemberF1(oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberF1"
    On Error GoTo ErrorHandler
    
    ' Transform the nozzle so that it behaves like a rigid body inside the equipment
    'm_oEqpCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 3
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for MemberF2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CMConstructMemberF2(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObj As Object)
    Const METHOD = "CMConstructMemberF2"
    On Error GoTo ErrorHandler

    ' Create Pipe Nozzle
    m_oEqpCADHelper.CreateDynamicPipeNozzleFromPH pMemberDescription, pResourceManager, pObj, 4
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructMemberF2(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructMemberF2"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsMemberF2(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsMemberF2"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateMemberF2(oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberF2"
    On Error GoTo ErrorHandler
    
    ' Transform the nozzle so that it behaves like a rigid body inside the equipment
    'm_oEqpCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 4
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for MemberF3
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CMConstructMemberF3(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObj As Object)
    Const METHOD = "CMConstructMemberF3"
    On Error GoTo ErrorHandler

    ' Create Pipe Nozzle
    m_oEqpCADHelper.CreateDynamicPipeNozzleFromPH pMemberDescription, pResourceManager, pObj, 5

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructMemberF3(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructMemberF3"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsMemberF3(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsMemberF3"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateMemberF3(oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberF3"
    On Error GoTo ErrorHandler
    
    ' Transform the nozzle so that it behaves like a rigid body inside the equipment
    'm_oEqpCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 5
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub HandleError(sModule As String, sMethod As String)
    
    If Not m_oEditErrors Is Nothing Then
        m_oEditErrors.AddFromErr Err, "", sMethod, sModule
    End If
    
    Err.Raise Err.Number, Err.Source & " " & sMethod, Err.Description, Err.HelpFile, Err.HelpContext
End Sub




