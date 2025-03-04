VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PumpDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:29:02 PM"

'******************************************************************
' Copyright (C) 2004, Intergraph Corporation. All rights reserved.
'
'File
'    PumpDef.cls
'
'Author
'       BBK
'
'Description
'   The symbol creates a Place Holder object for the position and orientation and other data of the nozzle.
'   This facilatate the creation of the nozzles in CAD and also useful in Solid Edge Equipments.
'
'Notes
'
'History:
'   23 Aug. 2004    BBK     Initial Creation
'
'   02.Jan.2005     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const MODULE = "PumpDef:"
Private Const m_ItemProgId As String = "SP3DSmartPump.PumpDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const GUID_CPEQUIPMENT = "{0D1E23F9-280A-4382-9370-9B0C5C363C88}"
Private Const GUID_CPEQUIPMENTCOMPONENT = "{F4FC2BB5-A507-4c59-B6E8-BC0AE3842A26}"

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
    oAD.AggregatorClsid = GUID_CPEQUIPMENT  'CPSmartEquipment
    oAD.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructPump"
    oAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructPump"
    oAD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsPump"
    oAD.SetCMRemoveInputs imsCOOKIE_ID_USS_LIB, "CMRemoveInputsPump"
    
    ' Add property to aggregator
    Set oPropertyDescriptions = oAD
    oPropertyDescriptions.RemoveAll
    oPropertyDescriptions.AddProperty "PumpProperties", 1, IID_IJDATTRIBUTES, "CMEvaluatePump", imsCOOKIE_ID_USS_LIB
    
    ' Remove all the previous member descriptions
    Set oMemberDescriptions = oSymbolDefinition
    oMemberDescriptions.RemoveAll
    
    ' Add new member(PipeNozzle1) to the definition
    Set oMemberDescription = oMemberDescriptions.AddMember("PipeNozzle1", 1, "CMConstructPipeNozzle1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsPipeNozzle1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructPipeNozzle1"
    ' Add properties for PipeNozzle1
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "PipeNozzle1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluatePipeNozzle1", imsCOOKIE_ID_USS_LIB

    ' Add new member(PipeNozzle2) to the definition
    Set oMemberDescription = oMemberDescriptions.AddMember("PipeNozzle2", 2, "CMConstructPipeNozzle2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsPipeNozzle2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructPipeNozzle2"
    ' Add properties for PipeNozzle2
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "PipeNozzle2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluatePipeNozzle2", imsCOOKIE_ID_USS_LIB
    
    ' Add new member(EqpFoundationPort) to the definition
    Set oMemberDescription = oMemberDescriptions.AddMember("EqpFoundationPort", 3, "CMConstructEqpFoundationPort", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEqpFoundationPort"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEqpFoundationPort"
    ' Add properties for PipeNozzle2
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "EqpFoundationPortProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEqpFoundationPort", imsCOOKIE_ID_USS_LIB
    
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

Public Sub CMConstructPump(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMConstructPump"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructPump(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructPump"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsPump(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMSetInputsPump"
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

Public Sub CMRemoveInputsPump(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMRemoveInputsPump"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluatePump(oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluatePump"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for PipeNozzle1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CMConstructPipeNozzle1(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObj As Object)
    Const METHOD = "CMConstructPipeNozzle1"
    On Error GoTo ErrorHandler
    
    m_oEqpCADHelper.CreateNozzlePHFromPH pMemberDescription, pResourceManager, pObj, 1

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructPipeNozzle1(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructPipeNozzle1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsPipeNozzle1(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsPipeNozzle1"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluatePipeNozzle1(oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluatePipeNozzle1"
    On Error GoTo ErrorHandler

    ' Transform the nozzle so that it behaves like a rigid body inside the equipment
    'm_oEqpCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 1
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for PipeNozzle2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CMConstructPipeNozzle2(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObj As Object)
    Const METHOD = "CMConstructPipeNozzle2"
    On Error GoTo ErrorHandler

    m_oEqpCADHelper.CreateNozzlePHFromPH pMemberDescription, pResourceManager, pObj, 2

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructPipeNozzle2(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructPipeNozzle2"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsPipeNozzle2(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInput2PipeNozzle2"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluatePipeNozzle2(oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluatePipeNozzle2"
    On Error GoTo ErrorHandler
    
    ' Transform the nozzle so that it behaves like a rigid body inside the equipment
    'm_oEqpCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 2
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for EqpFoundationPort
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CMConstructEqpFoundationPort(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObj As Object)
    Const METHOD = "CMConstructEqpFoundationPort"
    On Error GoTo ErrorHandler
   
    m_oEqpCADHelper.CreateNozzlePHFromPH pMemberDescription, pResourceManager, pObj, 3
   
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEqpFoundationPort(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructEqpFoundationPort"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEqpFoundationPort(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsEqpFoundationPort"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEqpFoundationPort(oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEqpFoundationPort"
    On Error GoTo ErrorHandler

    ' Transform the nozzle so that it behaves like a rigid body inside the equipment
    'm_oEqpCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 3
    
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


