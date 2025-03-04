VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VerShTExchDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"CAD"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:30:13 PM"
Attribute VB_Ext_KEY = "SP3DEqpCADTemplateVersion" ,"1.2"
'******************************************************************************
'   Copyright (C) 2007, Intergraph Corporation. All rights reserved.
'
'   File:      VerShTExchDef.cls
'   Author:    RUK
'   Creation Date:  Monday, Mar 5 2007
'
'   Description:
'   This is E310 Vertical Shell Tube Exchanger Equipment Assembly.
'   This Assembly has two DatumPoints with Orientations X towards Up and Y along North, and
'   One Circular FoundationPort
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "SP3DE_310VerShTubeExchAsm:VerShTExchDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors
Private m_avSymbolArrayOfInputs()   As Variant

Private m_oNorth        As IJDVector
Private m_oEast         As IJDVector
Private m_oElevation    As IJDVector

'Define Front End types
Private Const FrontEndType_A = 1
Private Const FrontEndType_B = 2
Private Const FrontEndType_C = 3
Private Const FrontEndType_D = 4
Private Const FrontEndType_N = 5
Private Const FrontEndType_Q = 6

'Define Rear End types
Private Const RearEndType_L = 1
Private Const RearEndType_M = 2
Private Const RearEndType_N = 3
Private Const RearEndType_P = 4
Private Const RearEndType_Q = 5
Private Const RearEndType_S = 6
Private Const RearEndType_T = 7
Private Const RearEndType_U = 8
Private Const RearEndType_W1 = 9
Private Const RearEndType_W2 = 10

Private m_dExchangerLength As Double
Private m_dExchangerDiameter As Double
Private m_dSupportDepth As Double
Private m_dSupportFromOrigin As Double
Private m_dFrontEndFlangeDia As Double
Private m_dFrontEndFlangeTk1 As Double
Private m_dFrontEndLength1 As Double
Private m_dFrontEndLength2 As Double
Private m_dFrontEndFlangeTk2 As Double
Private m_dFrontEndFlangeTk3 As Double
Private m_dRearEndFlangeDia As Double
Private m_dRearEndFlangeTk1 As Double
Private m_dRearEndLength As Double
Private m_dRearEndFlangeTk2 As Double
Private m_dRearEndFlangeTk3 As Double
Private m_dInsulationThickness As Double
Private m_dFrontEndType As Long
Private m_dRearEndType As Long
Private m_dChannelDiameter As Double
Private m_dFrontEndDiameter As Double
Private m_dRearEndDiameter As Double

Private m_bComputeEqpComp As Boolean
Private PI As Double

'Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt

Private Sub Class_Initialize()
    Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
    
    m_oEquipCADHelper.ProjectName = "SP3DE_310VerShTubeExchAsm"
    m_oEquipCADHelper.ClassName = "VerShTExchDef"
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
        
    'Add new member(TyACDN) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("TyACDN", 3, "CMConstructTyACDN", imsCOOKIE_ID_USS_LIB)
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
    Set oMemberDescription = oMemberDescriptions.AddMember("TyB", 4, "CMConstructTyB", imsCOOKIE_ID_USS_LIB)
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
    Set oMemberDescription = oMemberDescriptions.AddMember("TyFrontQ", 5, "CMConstructTyFrontQ", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsTyFrontQ"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTyFrontQ"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalTyFrontQ"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseTyFrontQ"

    'Add properties for (TyFrontQ)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "TyFrontQProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateTyFrontQ", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "TyFrontQGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryTyFrontQ", imsCOOKIE_ID_USS_LIB

    'Add new member(TyLNPW1) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("TyLNPW1", 6, "CMConstructTyLNPW1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsTyLNPW1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTyLNPW1"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalTyLNPW1"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseTyLNPW1"

    'Add properties for (TyLNPW1)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "TyLNPW1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateTyLNPW1", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "TyLNPW1GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryTyLNPW1", imsCOOKIE_ID_USS_LIB

    'Add new member(TyMSTUW2) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("TyMSTUW2", 7, "CMConstructTyMSTUW2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsTyMSTUW2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTyMSTUW2"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalTyMSTUW2"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseTyMSTUW2"

    'Add properties for (TyMSTUW2)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "TyMSTUW2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateTyMSTUW2", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "TyMSTUW2GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryTyMSTUW2", imsCOOKIE_ID_USS_LIB

    'Add new member(TyRearQ) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("TyRearQ", 8, "CMConstructTyRearQ", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsTyRearQ"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTyRearQ"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalTyRearQ"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseTyRearQ"

    'Add properties for (TyRearQ)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "TyRearQProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateTyRearQ", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "TyRearQGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryTyRearQ", imsCOOKIE_ID_USS_LIB
           
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
    On Error GoTo ErrorHandler
    
    Dim oMemberDescription As IJDMemberDescription
    Set oMemberDescription = m_oEquipCADHelper.GetMemberDescriptionFromChild(pIJDAttrs)
        
    Select Case oMemberDescription.Name
        Case "DP1", "DP2", "TyACDN", "TyB", "TyFrontQ", "TyLNPW1", "TyMSTUW2", "TyRearQ"
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
        Case "DP1", "DP2"
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
            Case "TyACDN", "TyB", "TyFrontQ", "TyLNPW1", "TyMSTUW2", "TyRearQ"
                For m = 1 To oAttrCollection.Count
                    Set oAttributeDescriptor = oAttrCollection.Item(m)
                    Select Case UCase(oAttributeDescriptor.InterfaceName)
                        Case "IJUAHEATEXCHANGERFRONTEND", "IJUAHEATEXCHANGERREAREND", "IJUAEXCHANGERCHANNEL", "IJUAHEATEXCHANGER"
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

'The code of this method is specific to each Equipment as it converts the array of inputs in a set of global variables
Private Sub GetDimensionsFromSymbolArray(SmartOccurrence As IJSmartOccurrence)
    Const METHOD = "GetDimensionsFromSymbolArray"
    On Error GoTo ErrorHandler
    
    m_avSymbolArrayOfInputs = m_oEquipCADHelper.GetSymbolArrayOfInputs(SmartOccurrence)

    'Inputs, from equipment symbol code
    m_dFrontEndType = m_avSymbolArrayOfInputs(2)
    m_dRearEndType = m_avSymbolArrayOfInputs(3)
    m_dExchangerDiameter = m_avSymbolArrayOfInputs(4)           'P2
    m_dExchangerLength = m_avSymbolArrayOfInputs(5)             'P1
    m_dSupportDepth = m_avSymbolArrayOfInputs(16)               'P13
    m_dSupportFromOrigin = m_avSymbolArrayOfInputs(17)          'P14
    m_dFrontEndFlangeDia = m_avSymbolArrayOfInputs(21)          'P30
    m_dFrontEndFlangeTk1 = m_avSymbolArrayOfInputs(22)          'P31
    m_dFrontEndFlangeTk2 = m_avSymbolArrayOfInputs(23)          'P34
    m_dFrontEndFlangeTk3 = m_avSymbolArrayOfInputs(24)          'P35
    m_dFrontEndLength1 = m_avSymbolArrayOfInputs(25)            'P32
    m_dFrontEndLength2 = m_avSymbolArrayOfInputs(26)            'P33
    m_dRearEndFlangeDia = m_avSymbolArrayOfInputs(27)           'P40
    m_dRearEndFlangeTk1 = m_avSymbolArrayOfInputs(28)           'P41
    m_dRearEndFlangeTk2 = m_avSymbolArrayOfInputs(29)           'P43
    m_dRearEndFlangeTk3 = m_avSymbolArrayOfInputs(30)           'P44
    m_dRearEndLength = m_avSymbolArrayOfInputs(31)              'P42
    m_dChannelDiameter = m_avSymbolArrayOfInputs(32)            'P43 of type M/S/T/U/W2 End
    m_dFrontEndDiameter = m_avSymbolArrayOfInputs(33)           'P33 of type Q End
    m_dRearEndDiameter = m_avSymbolArrayOfInputs(34)            'P43 of type Q End
    m_dInsulationThickness = m_avSymbolArrayOfInputs(35)
    
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
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oElevation, m_oNorth
    
    Set oPosition = Nothing
End Sub

Private Sub PositionAndOrientDP2(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    
    oPosition.Set 0, 0, -m_dSupportFromOrigin - m_dSupportDepth
    
    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oElevation, m_oNorth
    
    Set oPosition = Nothing
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

    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "FrontEndTypeACDN 01-EC", "TyACDN")
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oPosition As IJDPosition
    Set oPosition = New AutoMath.DPosition
    oPosition.Set 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformPositonOfEnds oDesignEquipment, pObject, oPosition, 0, 1, 0
    
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

    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "FrontEndTypeB 01-EC", "TyB")
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oPosition As IJDPosition
    Set oPosition = New AutoMath.DPosition
    oPosition.Set 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformPositonOfEnds oDesignEquipment, pObject, oPosition, 0, 1, 0

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

    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "FrontEndTypeQ 01-EC", "TyQ")
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oPosition As IJDPosition
    Set oPosition = New AutoMath.DPosition
    oPosition.Set 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformPositonOfEnds oDesignEquipment, pObject, oPosition, 0, 1, 0

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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for Variable Rear Ends
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructTyLNPW1(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructTyLNPW1"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "RearEndTypeLNPW1 01-EC", "TyLNPW1")
    GetDimensionsFromSymbolArray pMemberDescription.CAO

    Dim oPosition As IJDPosition
    Set oPosition = New AutoMath.DPosition
    oPosition.Set 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformPositonOfEnds oDesignEquipment, pObject, oPosition, 0, 1, 0
    Set oPosition = Nothing
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTyLNPW1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructTyLNPW1"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsTyLNPW1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsTyLNPW1"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence

    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAHeatExchanger").Item("ExchangerLength").Value = m_dExchangerLength
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength1").Value = m_dFrontEndLength1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndFlangeDia").Value = m_dRearEndFlangeDia
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndFlangeTk1").Value = m_dRearEndFlangeTk1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndLength").Value = m_dRearEndLength
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndFlangeTk2").Value = m_dRearEndFlangeTk2
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndFlangeTk3").Value = m_dRearEndFlangeTk3
    oAttribs.CollectionOfAttributes("IJUAEXCHANGERCHANNEL").Item("ChannelDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateTyLNPW1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateTyLNPW1"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryTyLNPW1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryTyLNPW1"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalTyLNPW1(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalTyLNPW1"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    GetDimensionsFromSymbolArray pMemberDesc.CAO

    If m_dRearEndType = RearEndType_L Or m_dRearEndType = RearEndType_N Or _
        m_dRearEndType = RearEndType_P Or m_dRearEndType = RearEndType_W1 Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseTyLNPW1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseTyLNPW1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConstructTyMSTUW2(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructTyMSTUW2"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "RearEndTypeMSTUW2 01-EC", "TyMSTUW2")
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oPosition As IJDPosition
    Set oPosition = New AutoMath.DPosition
    oPosition.Set 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformPositonOfEnds oDesignEquipment, pObject, oPosition, 0, 1, 0

    Set oPosition = Nothing
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTyMSTUW2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructTyMSTUW2"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsTyMSTUW2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsTyMSTUW2"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence

    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAHeatExchanger").Item("ExchangerLength").Value = m_dExchangerLength
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength1").Value = m_dFrontEndLength1

    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndFlangeDia").Value = m_dRearEndFlangeDia
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndFlangeTk1").Value = m_dRearEndFlangeTk1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndLength").Value = m_dRearEndLength
    oAttribs.CollectionOfAttributes("IJUAEXCHANGERCHANNEL").Item("ChannelDiameter").Value = m_dChannelDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateTyMSTUW2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateTyMSTUW2"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryTyMSTUW2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryTyMSTUW2"
    On Error GoTo ErrorHandler
    LogCalls METHOD

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalTyMSTUW2(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalTyMSTUW2"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    GetDimensionsFromSymbolArray pMemberDesc.CAO

    If m_dRearEndType = RearEndType_M Or m_dRearEndType = RearEndType_S Or _
        m_dRearEndType = RearEndType_T Or m_dRearEndType = RearEndType_U Or _
        m_dRearEndType = RearEndType_W2 Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseTyMSTUW2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseTyMSTUW2"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConstructTyRearQ(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructTyRearQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "RearEndTypeQ 01-EC", "TyQ")
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    Dim oPosition As IJDPosition
    Set oPosition = New AutoMath.DPosition
    oPosition.Set 0, 0, 0

    'To set the End according to the Occurnece Attributes
    TransformPositonOfEnds oDesignEquipment, pObject, oPosition, 0, 1, 0

    Set oPosition = Nothing
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTyRearQ(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructTyRearQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsTyRearQ(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsTyRearQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    GetDimensionsFromSymbolArray pMemberDesc.CAO
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence

    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAHeatExchanger").Item("ExchangerLength").Value = m_dExchangerLength
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength1").Value = m_dFrontEndLength1

    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndFlangeDia").Value = m_dRearEndFlangeDia
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndFlangeTk1").Value = m_dRearEndFlangeTk1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerRearEnd").Item("RearEndLength").Value = m_dRearEndLength
    oAttribs.CollectionOfAttributes("IJUAEXCHANGERCHANNEL").Item("RearEndDiameter").Value = m_dRearEndDiameter
    oAttribs.CollectionOfAttributes("IJUAEXCHANGERCHANNEL").Item("ChannelDiameter").Value = m_dExchangerDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateTyRearQ(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateTyRearQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryTyRearQ(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryTyRearQ"
    On Error GoTo ErrorHandler
    LogCalls METHOD
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalTyRearQ(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalTyRearQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    GetDimensionsFromSymbolArray pMemberDesc.CAO

    If m_dRearEndType = RearEndType_Q Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseTyRearQ(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseTyRearQ"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub TransformPositonOfEnds(Equipment As IJEquipment, Object As Object, oPosition As IJDPosition, _
 dRotAboutXaxis As Double, dRotAboutYaxis As Double, dRotAboutZaxis As Double)
    Const METHOD = "TransformPositonOfEnds"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oEqpMatrix As IJDT4x4
    Dim oTransform As IJDGeometry
    Dim oDirVector As New AutoMath.DVector

    If Not Object Is Nothing Then
        If TypeOf Object Is IJDGeometry Then
            Equipment.GetMatrix oEqpMatrix
            If Not oPosition Is Nothing Then
                oEqpMatrix.IndexValue(12) = oEqpMatrix.IndexValue(12) + oPosition.x
                oEqpMatrix.IndexValue(13) = oEqpMatrix.IndexValue(13) + oPosition.y
                oEqpMatrix.IndexValue(14) = oEqpMatrix.IndexValue(14) + oPosition.z
            End If
            Set oTransform = Object
            'Apply Rotation
            If Not dRotAboutXaxis = 0 Then
                oDirVector.Set 1, 0, 0
                oEqpMatrix.Rotate PI / 2, oDirVector
            End If
            If Not dRotAboutYaxis = 0 Then
                oDirVector.Set 0, 1, 0
                oEqpMatrix.Rotate PI / 2, oDirVector
            End If
            If Not dRotAboutZaxis = 0 Then
                oDirVector.Set 0, 0, 1
                oEqpMatrix.Rotate PI / 2, oDirVector
            End If
            oTransform.DTransform oEqpMatrix
        End If
    End If

    Set oTransform = Nothing
    Set oEqpMatrix = Nothing
    Set oDirVector = Nothing
    
    Exit Sub
ErrorHandler:
    Set oTransform = Nothing
    Set oEqpMatrix = Nothing
    Set oDirVector = Nothing
    HandleError MODULE, METHOD
End Sub

